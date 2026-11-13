"""
Verificare automata a starii posturilor de contorizare trafic (ADR/Lantronix/Wiznet)
======================================================================================

Bazat pe analiza log-urilor Portmon capturate pe masina virtuala XP, exista 3 stari
posibile pentru un post, distinse astfel:

  1. RETEA_CAZUTA  -> conexiunea TCP catre IP:port nu se stabileste deloc
                       (echivalent cu IRP_MJ_CREATE esuat in Portmon)
                       => router fara curent / SIM fara semnal-credit / Wiznet mort

  2. CONTOR_CAZUT   -> conexiunea TCP se stabileste, se trimite secventa de
                       "trezire" + comanda de interogare (#I), dar nu vine niciun
                       raspuns valid inapoi in timpul alocat
                       => router + Wiznet functioneaza, dar contorul ADR nu (baterie
                          descarcata, alimentare intrerupta, echipament ars)

  3. OK             -> conexiunea TCP se stabileste SI se primeste un raspuns valid
                       (incepe cu prefixul "#I") la comanda de interogare.

Bytes-ii de mai jos (WAKE_SEQUENCE, IDENTIFY_CMD, EXPECTED_RESPONSE_PREFIX) au fost
extrasi din capturile Portmon reale (snagov.LOG - post functional) si confirmati ca
fiind identici (deci comanda e generica, nu are adresa contorului codificata) in
nu_merge_contorul.LOG, unde aceeasi comanda e trimisa repetat fara sa vina niciodata
raspuns - exact semnatura "CONTOR_CAZUT" descrisa mai sus.
"""

import os
import re
import sys
import glob
import socket
import threading
import queue
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Foloseste noua baza SQLite de identificare contoare (database.py / contoare.db),
# via singleton-ul get_contoare_db(). database.py si config.py trebuie sa fie in
# acelasi folder (sau in PYTHONPATH) cand rulezi acest script.
from database import get_contoare_db

# ---------------------------------------------------------------------------
# CONFIGURARE
# ---------------------------------------------------------------------------

PORT_FIX = 1985              # portul e mereu acelasi pentru toate posturile

# Verificam doar contoarele ADR (nu si VEKs4). Comparatia se face fara sa conteze
# majusculele/spatiile, dar in raport ramane afisata valoarea originala din baza.
TIPURI_PERMISE = {"adr 3000", "adr 2000", "adr sabre"}

TCP_CONNECT_TIMEOUT = 5.0   # secunde - cat asteptam sa se stabileasca conexiunea TCP

# Unele contoare NU raspund la prima comanda #I trimisa - au nevoie sa fie
# "batute la usa" de cateva ori inainte sa reactioneze (vezi post_ok.LOG: raspunsul
# a venit abia la a 3-a retrimitere a comenzii). De aceea retrimitem periodic in
# loc sa asteptam pasiv un singur timeout lung.
PROTOCOL_RETRY_COUNT = 6      # de cate ori retrimitem WAKE+IDENTIFY daca nu vine raspuns
PROTOCOL_RETRY_INTERVAL = 10.0  # secunde intre retrimiteri (buget total ~= COUNT*INTERVAL)

MAX_WORKERS = 20            # cate posturi verificam in paralel

# Verificarea zilnica: un post care nu e OK la primul tur este reverificat (doar el)
# in inca INCERCARI_ZILNICE-1 ture, cu o pauza intre ture. Se opreste la primul OK.
INCERCARI_ZILNICE = 3
PAUZA_INTRE_TURE = 60.0     # secunde intre ture (pauza e comuna, nu blocheaza cate un thread)

# Bytes reali, extrasi din capturile Portmon (snagov.LOG / nu_merge_contorul.LOG)
WAKE_SEQUENCE = bytes.fromhex("00FF00FF00FF00FF00FF00FF00FF00FF00FF")  # 18 bytes
IDENTIFY_CMD = bytes.fromhex("234904000100010C0100C0A0")               # 12 bytes: #I...
EXPECTED_RESPONSE_PREFIX = bytes.fromhex("2349")                       # "#I"

# Folderul unde se salveaza rapoartele. Incercam sa preluam CENTRAL_FILE_FOLDER
# din config.py (acelasi folosit de contoare_db.py); daca nu exista, cadem pe
# radacina fixa de mai jos.
try:
    from config import CENTRAL_FILE_FOLDER
    OUTPUT_FOLDER = os.path.join(CENTRAL_FILE_FOLDER, "Verificare teletransmisie")
except ImportError:
    OUTPUT_FOLDER = r"L:\BIDMRCT\datePEEK\Verificare teletransmisie"

FONT_NAME = "Times New Roman"

# Ordinea dorita de afisare a DRDP-urilor in raport (Bucuresti primul, apoi in
# aceasta ordine). Comparatia ignora diacriticele si majusculele.
DRDP_ORDINE_NUME = ["Bucuresti", "Craiova", "Timisoara", "Cluj", "Brasov",
                     "Iasi", "Constanta", "Buzau"]


def normalizeaza(text: str) -> str:
    """Elimina diacriticele si normalizeaza la litere mici, pentru comparatii sigure."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return text.strip().lower()


DRDP_ORDINE = {normalizeaza(nume): idx for idx, nume in enumerate(DRDP_ORDINE_NUME)}

# ---------------------------------------------------------------------------
# LOGICA DE VERIFICARE
# ---------------------------------------------------------------------------


@dataclass
class RezultatPost:
    post_id: str
    localitate: str
    drdp: str
    tip: str
    ip: str
    port: int
    status: str          # "RETEA_CAZUTA" | "CONTOR_CAZUT" | "OK" | "EROARE"
    detaliu: str = ""
    verificat_la: str = ""   # ora exacta la care s-a terminat verificarea acestui post
    incercari: int = 1       # cate ture a fost nevoie (1 = OK/esec din prima)


def citeste_posturi() -> list[tuple[str, str, str, str, str, int]]:
    """
    Citeste lista de posturi (post_id, localitate, drdp, tip, ip, port) din
    contoare.db, via database.get_contoare_db().get_all(), care returneaza un
    dict de forma:
        { "2168": {"Drum": "DN1", "Localitate": "...", "Tip": "...",
                    "IP": "...", "DRDP": "...", ...}, ... }
    Sar peste posturile fara IP completat (necunoscute/neinstalate) si peste
    cele care nu sunt de tip ADR 3000 / ADR 2000 / ADR Sabre (fara VEKs4).
    """
    db = get_contoare_db().get_all()
    posturi = []
    for ct_id, date in sorted(db.items()):
        ip = (date.get("IP") or "").strip()
        if not ip:
            continue
        tip = (date.get("Tip") or "").strip()
        if tip.lower() not in TIPURI_PERMISE:
            continue
        localitate = (date.get("Localitate") or "").strip()
        drdp = (date.get("DRDP") or "").strip()
        posturi.append((ct_id, localitate, drdp, tip, ip, PORT_FIX))
    return posturi


def verifica_post(post_id: str, localitate: str, drdp: str, tip: str,
                   ip: str, port: int) -> RezultatPost:
    """Verifica un singur post si returneaza starea lui pe cele 3 nivele."""
    sock = None
    status = "EROARE"
    detaliu = ""
    try:
        # --- Tier 1: conexiune TCP (echivalent IRP_MJ_CREATE) ---
        try:
            sock = socket.create_connection((ip, port), timeout=TCP_CONNECT_TIMEOUT)
        except (socket.timeout, ConnectionRefusedError, OSError) as e:
            status, detaliu = "RETEA_CAZUTA", str(e)
        else:
            # --- Tier 2: handshake protocol ADR (echivalent secventei #I din log) ---
            # Retrimitem comanda de mai multe ori - unele contoare (cu link mai slab)
            # nu raspund la prima incercare, dar raspund la a 2-a/3-a (vezi post_ok.LOG).
            try:
                raspuns = b""
                for incercare in range(1, PROTOCOL_RETRY_COUNT + 1):
                    sock.sendall(WAKE_SEQUENCE)
                    sock.sendall(IDENTIFY_CMD)

                    deadline = time.monotonic() + PROTOCOL_RETRY_INTERVAL
                    while time.monotonic() < deadline:
                        sock.settimeout(max(0.05, deadline - time.monotonic()))
                        try:
                            chunk = sock.recv(256)
                        except socket.timeout:
                            break
                        if not chunk:
                            break
                        raspuns += chunk
                        if EXPECTED_RESPONSE_PREFIX in raspuns:
                            break

                    if EXPECTED_RESPONSE_PREFIX in raspuns:
                        break

                if EXPECTED_RESPONSE_PREFIX in raspuns:
                    status, detaliu = "OK", ""
                else:
                    status = "CONTOR_CAZUT"
                    detaliu = (f"niciun byte primit dupa {PROTOCOL_RETRY_COUNT} incercari (timeout)"
                               if not raspuns
                               else f"raspuns neasteptat dupa {PROTOCOL_RETRY_COUNT} incercari: {raspuns!r}")
            except (socket.timeout, OSError) as e:
                status, detaliu = "CONTOR_CAZUT", str(e)

    except Exception as e:  # noqa: BLE001 - vrem sa prindem orice si sa raportam
        status, detaliu = "EROARE", str(e)
    finally:
        if sock is not None:
            sock.close()

    verificat_la = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    return RezultatPost(post_id, localitate, drdp, tip, ip, port, status, detaliu, verificat_la)


def _tur_verificare(posturi: list[tuple[str, str, str, str, str, int]]) -> list[RezultatPost]:
    """Un singur tur: verifica posturile date, in paralel, folosind un pool simplu de threaduri."""
    q: queue.Queue = queue.Queue()
    for p in posturi:
        q.put(p)

    rezultate: list[RezultatPost] = []
    lock = threading.Lock()

    def worker():
        while True:
            try:
                post_id, localitate, drdp, tip, ip, port = q.get_nowait()
            except queue.Empty:
                return
            rez = verifica_post(post_id, localitate, drdp, tip, ip, port)
            with lock:
                rezultate.append(rez)
                # afisare live in terminal, pe masura ce vin rezultatele. Un esec de
                # afisare (ex. encoding la redirectare in fisier) NU trebuie sa opreasca threadul.
                try:
                    print(f"[{rez.status:13s}] {rez.post_id} - {rez.localitate} "
                          f"(DRDP {rez.drdp}, {rez.tip})  ({rez.ip}:{rez.port})", flush=True)
                except Exception:  # noqa: BLE001
                    pass
            q.task_done()

    threads = [threading.Thread(target=worker) for _ in range(min(MAX_WORKERS, len(posturi)))]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    return rezultate


def verifica_toate_posturile(posturi: list[tuple[str, str, str, str, str, int]],
                             incercari: int = INCERCARI_ZILNICE,
                             pauza: float = PAUZA_INTRE_TURE) -> list[RezultatPost]:
    """Verifica posturile in pana la `incercari` ture. La fiecare tur se reiau doar
    posturile care nu au fost OK; un post OK iese din joc. Rezultatul final al unui
    post = cel din ultimul tur in care a fost verificat (OK daca a raspuns macar o data)."""
    final: dict[str, RezultatPost] = {}
    de_verificat = list(posturi)
    for tur in range(1, max(1, incercari) + 1):
        if not de_verificat:
            break
        if tur > 1:
            print(f"--- Tur {tur}/{incercari}: reverific {len(de_verificat)} posturi "
                  f"(dupa {pauza:.0f}s pauza) ---")
            time.sleep(pauza)
        for rez in _tur_verificare(de_verificat):
            rez.incercari = tur
            final[rez.post_id] = rez
        de_verificat = [p for p in de_verificat
                        if p[0] not in final or final[p[0]].status != "OK"]
    # siguranta: orice post care n-a primit niciun rezultat e marcat EROARE, nu dispare
    for post_id, localitate, drdp, tip, ip, port in posturi:
        if post_id not in final:
            final[post_id] = RezultatPost(post_id, localitate, drdp, tip, ip, port, "EROARE",
                                          "verificarea nu a produs niciun rezultat",
                                          datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
                                          incercari)
    return list(final.values())


# ---------------------------------------------------------------------------
# ISTORIC ZILNIC IN contoare.db (folosit de raport_mod_functionare.py)
# ---------------------------------------------------------------------------

def salveaza_verificari(rezultate: list[RezultatPost], data: date | None = None) -> int:
    """Salveaza rezultatele verificarii pentru ziua `data` (implicit azi) in contoare.db.
    Daca ziua a mai fost verificata, randurile aceluiasi post sunt inlocuite."""
    data = data or date.today()
    rows = [(r.post_id, r.status, r.detaliu, r.incercari, r.verificat_la) for r in rezultate]
    return get_contoare_db().save_verificari(data.strftime("%Y-%m-%d"), rows)


def citeste_verificari(data: date) -> dict[str, dict]:
    """{post: {'status', 'detaliu', 'incercari', 'verificat_la'}} pentru ziua `data`."""
    return get_contoare_db().get_verificari(data.strftime("%Y-%m-%d"))


# ---------------------------------------------------------------------------
# RAPORT EXCEL
# ---------------------------------------------------------------------------

CULORI_STATUS = {
    "OK": "C6EFCE",
    "CONTOR_CAZUT": "FFC7CE",
    "RETEA_CAZUTA": "FFEB9C",
    "EROARE": "D9D9D9",
}

ETICHETE_STATUS = {
    "OK": "OK",
    "CONTOR_CAZUT": "Contor cazut (baterie/alimentare)",
    "RETEA_CAZUTA": "Retea cazuta (router/SIM)",
    "EROARE": "Eroare neasteptata",
}


def gaseste_raport_anterior(folder: str) -> str | None:
    """Returneaza calea celui mai recent raport_posturi_*.xlsx deja existent in folder."""
    if not os.path.isdir(folder):
        return None
    fisiere = glob.glob(os.path.join(folder, "raport_posturi_*.xlsx"))
    if not fisiere:
        return None
    fisiere.sort(key=os.path.getmtime)   # cel mai recent la final
    return fisiere[-1]


def citeste_statusuri_anterioare(cale_fisier: str) -> tuple[dict[str, str], str]:
    """
    Citeste dintr-un raport anterior maparea post_id -> status (eticheta text),
    plus data/ora la care a fost generat acel raport. Detecteaza randul de antet
    cautand celula "Post" (functioneaza indiferent de rand, ca sa nu conteze daca
    fisierul anterior are sau nu rand de rezumat deasupra).
    """
    statusuri: dict[str, str] = {}
    data_raport = "necunoscuta"
    try:
        wb = load_workbook(cale_fisier, data_only=True)
        ws = wb.active

        header_row = None
        for r in range(1, min(ws.max_row, 5) + 1):
            if str(ws.cell(r, 1).value).strip() == "Post":
                header_row = r
                break

        if header_row is not None:
            headers = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(header_row, c).value
                if v is not None:
                    headers[str(v).strip()] = c

            col_post = headers.get("Post")
            col_status = headers.get("Status")
            if col_post is not None and col_status is not None:
                for r in range(header_row + 1, ws.max_row + 1):
                    post_val = ws.cell(r, col_post).value
                    if post_val is None:
                        continue
                    status_val = ws.cell(r, col_status).value
                    statusuri[str(post_val).strip()] = str(status_val).strip() if status_val is not None else ""

        wb.close()

        nume = os.path.basename(cale_fisier)
        m = re.search(r"raport_posturi_(\d{2}\.\d{2}\.\d{4})_(\d{2})-(\d{2})-(\d{2})", nume)
        if m:
            data_raport = f"{m.group(1)} {m.group(2)}:{m.group(3)}:{m.group(4)}"
        else:
            data_raport = datetime.fromtimestamp(os.path.getmtime(cale_fisier)).strftime("%d.%m.%Y %H:%M:%S")
    except Exception as e:
        print(f"[WARN] nu am putut citi raportul anterior ({cale_fisier}): {e}")

    return statusuri, data_raport


def adauga_sheet_legenda(wb: Workbook) -> None:
    """Adauga un sheet separat cu explicatii ale statusurilor si ale mesajelor
    posibile din coloana Detaliu, ca sa se stie ce inseamna daca apar vreodata."""
    ws = wb.create_sheet("Legenda")
    ws.sheet_view.showGridLines = False

    font_titlu = Font(name=FONT_NAME, bold=True, size=13)
    font_subtitlu = Font(name=FONT_NAME, bold=True, size=11)
    font_antet = Font(name=FONT_NAME, bold=True, color="FFFFFF")
    font_normal = Font(name=FONT_NAME)
    fill_antet = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    wrap = Alignment(wrap_text=True, vertical="top")

    r = 1
    ws.cell(r, 1, "Semnificatia statusurilor si a mesajelor din coloana Detaliu").font = font_titlu
    r += 2

    # --- tabelul statusurilor ---
    ws.cell(r, 1, "Status").font = font_antet
    ws.cell(r, 2, "Ce inseamna").font = font_antet
    for c in (1, 2):
        ws.cell(r, c).fill = fill_antet
    r += 1
    statusuri_expl = [
        ("OK",
         "Contorul a raspuns corect (raspuns care incepe cu #I) la comanda de interogare - "
         "totul functioneaza, de la router pana la contor."),
        ("Contor cazut (baterie/alimentare)",
         "Conexiunea TCP catre router/placuta Wiznet reuseste, deci acea parte functioneaza, "
         "dar contorul ADR nu raspunde deloc (sau raspunde altceva decat asteptat) dupa toate "
         "incercarile - cel mai probabil baterie descarcata, alimentare intrerupta sau "
         "echipament ars."),
        ("Retea cazuta (router/SIM)",
         "Nu s-a putut stabili deloc conexiunea TCP catre IP:port-ul postului - router fara "
         "curent, cartela SIM fara semnal/credit, sau placuta Wiznet moarta."),
        ("Eroare neasteptata",
         "A aparut o eroare Python neprevazuta in timpul verificarii acelui post (nu se "
         "incadreaza la niciuna din categoriile de mai sus) - in mod normal nu ar trebui sa "
         "apara; daca apare, mesajul din Detaliu ajuta la diagnosticare."),
    ]
    for status, explicatie in statusuri_expl:
        ws.cell(r, 1, status).font = font_normal
        ws.cell(r, 2, explicatie).font = font_normal
        ws.cell(r, 1).alignment = wrap
        ws.cell(r, 2).alignment = wrap
        r += 1

    r += 2

    # --- tabelul mesajelor din Detaliu, grupate pe status ---
    ws.cell(r, 1, "Mesaje posibile in coloana Detaliu, pe status").font = font_subtitlu
    r += 1

    mesaje_grupate = [
        ("Retea cazuta (router/SIM)", [
            ("timed out",
             "Nu a raspuns nimic in timpul alocat conectarii (5 secunde) - SIM fara semnal sau router oprit."),
            ("[WinError 10061] ...refused / Connection refused",
             "Ceva raspunde pe adresa respectiva, dar refuza explicit conexiunea pe portul 1985 - "
             "posibil un echipament configurat gresit sau alt serviciu pe acel port."),
            ("[WinError 10060] ...",
             "Varianta specifica Windows a timeout-ului la conectare."),
            ("No route to host",
             "Problema de rutare in retea, nu doar semnal slab - de investigat la nivel de retea/VPN."),
        ]),
        ("Contor cazut (baterie/alimentare)", [
            (f"niciun byte primit dupa {PROTOCOL_RETRY_COUNT} incercari (timeout)",
             "Conexiunea catre router/Wiznet e OK, dar contorul insusi nu a raspuns deloc, desi "
             "comanda a fost retrimisa de mai multe ori."),
            ("raspuns neasteptat dupa N incercari: b'...'",
             "A venit un raspuns, dar nu incepe cu #I - posibil alt echipament pe linie, protocol "
             "diferit, sau zgomot pe legatura. Diferit de tacere totala, merita investigat separat."),
            ("[WinError 10054] ...forcibly closed / Connection reset",
             "Conexiunea a fost inchisa brusc de partea cealalta chiar in timpul comunicarii "
             "(nu la conectarea initiala)."),
            ("[Errno 32] Broken pipe",
             "Legatura s-a intrerupt exact in timp ce se trimiteau date catre post."),
        ]),
    ]

    for nume_status, mesaje in mesaje_grupate:
        ws.cell(r, 1, nume_status).font = font_subtitlu
        r += 1
        ws.cell(r, 1, "Mesaj in Detaliu").font = font_antet
        ws.cell(r, 2, "Semnificatie").font = font_antet
        for c in (1, 2):
            ws.cell(r, c).fill = fill_antet
        r += 1
        for mesaj, explicatie in mesaje:
            ws.cell(r, 1, mesaj).font = font_normal
            ws.cell(r, 2, explicatie).font = font_normal
            ws.cell(r, 1).alignment = wrap
            ws.cell(r, 2).alignment = wrap
            r += 1
        r += 1

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 90


def genereaza_raport(rezultate: list[RezultatPost], cale_fisier: str) -> None:
    folder = os.path.dirname(cale_fisier)
    raport_anterior = gaseste_raport_anterior(folder)
    if raport_anterior:
        statusuri_anterioare, data_raport_anterior = citeste_statusuri_anterioare(raport_anterior)
    else:
        statusuri_anterioare, data_raport_anterior = {}, "nicio verificare anterioara gasita"

    wb = Workbook()
    ws = wb.active
    ws.title = "Stare posturi"

    font_antet = Font(name=FONT_NAME, bold=True)
    font_date = Font(name=FONT_NAME)
    font_rezumat = Font(name=FONT_NAME, bold=True, size=12)

    n_ok = sum(1 for r in rezultate if r.status == "OK")
    n_contor = sum(1 for r in rezultate if r.status == "CONTOR_CAZUT")
    n_retea = sum(1 for r in rezultate if r.status == "RETEA_CAZUTA")
    n_eroare = sum(1 for r in rezultate if r.status == "EROARE")

    antet = ["Post", "Localitate", "DRDP", "Tip", "IP", "Status", "Detaliu", "Verificat la",
             f"Status verificare anterioara ({data_raport_anterior})"]

    # --- rand 1: rezumatul, acelasi format ca in consola ---
    ws.append([f"OK: {n_ok} | Contor cazut: {n_contor} | Retea cazuta: {n_retea} | Erori: {n_eroare}"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(antet))
    ws.cell(1, 1).font = font_rezumat
    ws.cell(1, 1).alignment = Alignment(horizontal="center")

    # --- rand 2: antetul coloanelor ---
    ws.append(antet)
    for cell in ws[2]:
        cell.font = font_antet

    # ordonam: intai cele problematice (RETEA_CAZUTA, CONTOR_CAZUT, EROARE, OK),
    # apoi in cadrul fiecarui status dupa DRDP (Bucuresti, Craiova, Timisoara, Cluj,
    # Brasov, Iasi, Constanta, Buzau - restul la final)
    ordine_status = {"RETEA_CAZUTA": 0, "CONTOR_CAZUT": 1, "EROARE": 2, "OK": 3}
    cheie_sortare = lambda x: (ordine_status.get(x.status, 9),
                                DRDP_ORDINE.get(normalizeaza(x.drdp), 999), x.post_id)
    for r in sorted(rezultate, key=cheie_sortare):
        status_anterior = statusuri_anterioare.get(r.post_id, "")
        ws.append([r.post_id, r.localitate, r.drdp, r.tip, r.ip,
                   ETICHETE_STATUS.get(r.status, r.status), r.detaliu, r.verificat_la,
                   status_anterior])
        fill = PatternFill(start_color=CULORI_STATUS.get(r.status, "FFFFFF"),
                            end_color=CULORI_STATUS.get(r.status, "FFFFFF"),
                            fill_type="solid")
        for cell in ws[ws.max_row]:
            cell.fill = fill
            cell.font = font_date

    latimi = [10, 20, 10, 14, 16, 32, 40, 20, len(antet[-1]) + 2]
    for col, width in zip("ABCDEFGHI", latimi):
        ws.column_dimensions[col].width = width

    ultimul_rand = ws.max_row
    ws.freeze_panes = "A3"   # ramane rezumatul + antetul fixe cand se face scroll
    ws.auto_filter.ref = f"A2:{get_column_letter(len(antet))}{ultimul_rand}"

    adauga_sheet_legenda(wb)

    os.makedirs(folder, exist_ok=True)
    wb.save(cale_fisier)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Cand iesirea e redirectata in fisier (Task Scheduler / .bat), Windows foloseste
    # cp1252, care nu poate scrie ă, ș, ț etc. Fortam UTF-8 (cu inlocuire la nevoie).
    for _flux in (sys.stdout, sys.stderr):
        try:
            _flux.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass

    posturi = citeste_posturi()
    print(f"Verific {len(posturi)} posturi (ADR 3000 / ADR 2000 / ADR Sabre)...")

    rezultate = verifica_toate_posturile(posturi)

    n_ok = sum(1 for r in rezultate if r.status == "OK")
    n_contor = sum(1 for r in rezultate if r.status == "CONTOR_CAZUT")
    n_retea = sum(1 for r in rezultate if r.status == "RETEA_CAZUTA")
    n_eroare = sum(1 for r in rezultate if r.status == "EROARE")

    print(f"OK: {n_ok} | Contor cazut: {n_contor} | Retea cazuta: {n_retea} | Erori: {n_eroare}")

    # Protectie: daca TOATE posturile par "retea cazuta", cel mai probabil calculatorul
    # n-are acces la reteaua contoarelor (VPN neconectat etc.) - nu salvam rezultate false.
    retea_indisponibila = (len(rezultate) >= 10 and n_ok == 0
                           and n_retea >= 0.95 * len(rezultate))
    if retea_indisponibila:
        print("ATENTIE: aproape toate posturile apar ca 'Retea cazuta' - probabil lipsa acces "
              "la retea/VPN de pe acest calculator. Rezultatele NU au fost salvate in istoric.")
    else:
        try:
            n_salvate = salveaza_verificari(rezultate)
            print(f"Istoric zilnic: {n_salvate} posturi salvate in contoare.db (verificare_zilnica)")
        except Exception as e:  # noqa: BLE001 - raportul Excel se genereaza oricum
            print(f"[WARN] Nu am putut salva istoricul zilnic in contoare.db: {e}")

    nume_fisier = datetime.now().strftime("raport_posturi_%d.%m.%Y_%H-%M-%S.xlsx")
    cale_completa = os.path.join(OUTPUT_FOLDER, nume_fisier)
    genereaza_raport(rezultate, cale_completa)
    print(f"Raport salvat in {cale_completa}")
