# =============================================================================
# raport_mod_functionare.py — Raport Excel: modul de funcționare al contoarelor
# active, pe benzi, pentru o dată aleasă.
# =============================================================================
#
# Utilizare:
#   python raport_mod_functionare.py 28.10.2026
#   python raport_mod_functionare.py 28.10.2026 "D:\Rapoarte\raport.xlsx"
#   Run din PyCharm (fără argumente): se deschide o fereastră în care alegi data
#
# Sau din alt modul:
#   from raport_mod_functionare import genereaza_raport
#   cale = genereaza_raport(datetime.date(2026, 10, 28))
#
# Starea fiecărei benzi (aceeași logică ca în foaia „Date prelucrate”):
#   Defect        : total vehicule în zi == 0
#   Totalizator   : Clasa_15 / Total bandă > PRAG_TOT (30%)   -> senzor piezoelectric
#                   SAU Clasa_2 < Clasa_4                      -> buclă inductivă de ieșire
#   Clasificator  : altfel
#   În plus, la contoare cu 2 benzi (B1-B2) sau 4 benzi (B1-B4, B2-B3): banda mai
#   mică, cu trafic > 0, dacă perechea ei are >= 2x mai mult -> „nu înregistrează
#   tot traficul” + „Necesita verificare bucla de intrare”.
# =============================================================================

import os
import sys
import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import CENTRAL_FILE_FOLDER
import sqlite3
from database import get_traffic_db, get_contoare_db, CONTOARE_DB
from verificare_posturi import citeste_verificari, ETICHETE_STATUS

PRAG_TOT = 0.3           # prag valabil DOAR pentru acest raport (Clasa_15 / Total)
MAX_BENZI = 6
FOLDER_RAPOARTE = "Rapoarte"   # subfolder in CENTRAL_FILE_FOLDER unde se salveaza raportul
OMITE_REDENUMITE = True   # True = nu apar posturile cu cod vechi (redenumite) din contor_alias

HEADERS = [
    "DRDP", "Drum", "Pozitie km.", "Post", "Localitatea", "IP",
    "Mod de functionare", "Piese/materiale necesare/recomandari", "Probleme",
]
COL_WIDTHS = [14, 10, 14, 10, 24, 16, 38, 71, 12]   # col. H (Piese/recomandari) 71 ≈ 500 px


# ── Logică stare bandă ───────────────────────────────────────────────────────

def _stare_banda(total, cls15, cls2, cls4):
    """Returnează (stare, piese_suplimentare):
    stare: 'D' = defect, 'T' = totalizator, 'C' = clasificator.
    piese: set cu 'piezo' și/sau 'bucla_iesire' (doar pentru totalizator)."""
    if total == 0:
        return "D", set()
    piese = set()
    if cls15 / total > PRAG_TOT:
        piese.add("piezo")
    if cls2 < cls4:
        piese.add("bucla_iesire")
    return ("T" if piese else "C"), piese


_MOD_TXT = {
    "C": "B{b} functioneaza corespunzator",
    "T": "B{b} functioneaza totalizator",
    "D": "B{b} nu inregistreaza trafic",
}
_MOD_INCOMPLET = "B{b} nu inregistreaza tot traficul"
_PIESE_INCOMPLET = "Necesita verificare bucla de intrare B{b}"
_PIESE_TXT = {
    "D":            "Necesita inlocuire bucle inductive B{b}",
    "piezo":        "Senzor piezoelectric pentru B{b}",
    "bucla_iesire": "Necesita inlocuire bucla inductiva de iesire B{b}",
}


# ── Posturi redenumite ───────────────────────────────────────────────────────

def _coduri_redenumite():
    """Setul codurilor VECHI din tabelul contor_alias (cod_vechi -> cod_nou)."""
    try:
        conn = sqlite3.connect(CONTOARE_DB, timeout=10)
        try:
            return {str(r[0]).strip() for r in
                    conn.execute("SELECT cod_vechi FROM contor_alias").fetchall()}
        finally:
            conn.close()
    except Exception:
        return set()      # tabelul nu exista inca -> nimic redenumit


# ── Motivul lipsei datelor (din verificarea zilnica a posturilor) ─────────────

_RECOMANDARE_STATUS = {
    "RETEA_CAZUTA": "Verificare router / SIM  / Wiznet",
    "CONTOR_CAZUT": "Verificare baterie / alimentare contor",
    "OK":           "Verificati descarcarea / prelucrarea datelor din aceasta zi",
    "EROARE":       "Eroare la verificare - reluati verificarea",
}


def _motiv_fara_date(ver, info):
    """Linia 'Motiv: ...' pentru un contor fara date, pe baza verificarii din ziua raportului.
    `ver` = {'status','detaliu','incercari','verificat_la'} sau None; `info` = randul din contoare.db."""
    if ver is None:
        ip = str(info.get("IP", "") or "").strip()
        tip = str(info.get("Tip", "") or "").strip()
        if not ip:
            return "Motiv: neverificabil (fara IP in baza de date)"
        if tip.lower() not in {"adr 3000", "adr 2000", "adr sabre"}:
            return f"Motiv: neverificabil (tip {tip or 'necunoscut'})"
        return "Motiv: verificarea zilnica nu a fost efectuata in aceasta zi"
    if ver["status"] == "OK":
        txt = "Motiv: contorul raspunde la verificare, dar nu are date inregistrate"
    else:
        txt = f"Motiv: {ETICHETE_STATUS.get(ver['status'], ver['status'])}"
    if ver["detaliu"]:
        txt += f" - {ver['detaliu']}"
    nr = ver.get("incercari") or 1
    txt += f"\nVerificat {ver['verificat_la']} (incercari: {nr})"
    rec = _RECOMANDARE_STATUS.get(ver["status"])
    if rec:
        txt += f"\nRecomandare: {rec}"
    return txt


# ── Benzi pereche (doar contoare cu 2 sau 4 benzi) ───────────────────────────

PERECHI = {2: [(1, 2)], 4: [(1, 4), (2, 3)]}
RAPORT_PERECHE = 2          # banda mai mica de cel putin de 2 ori fata de perechea ei


def _benzi_incomplete(n, tot):
    """Benzile care nu inregistreaza tot traficul: banda mai mica dintr-o pereche,
    cu trafic > 0, cand perechea ei are cel putin RAPORT_PERECHE x mai mult."""
    rez = set()
    for a, b in PERECHI.get(n, []):
        ta, tb = tot[a], tot[b]
        mic, mare = (a, b) if ta < tb else (b, a)
        if tot[mic] > 0 and tot[mare] >= RAPORT_PERECHE * tot[mic]:
            rez.add(mic)
    return rez


# ── Citire date zilnice din trafic.db ────────────────────────────────────────

def _citeste_zi(data):
    """Returnează {contor: {'n': n_benzi, 'tot': {b: int}, 'c15': {b: int}}}
    pentru ziua `data`, agregat din inregistrari_orare."""
    sel = ["contor", "MAX(n_benzi) AS n"]
    for b in range(1, MAX_BENZI + 1):
        sel.append(f"COALESCE(SUM(total_b{b}), 0) AS t{b}")
        sel.append(f"COALESCE(SUM(b{b}_clasa_15), 0) AS c{b}")
        sel.append(f"COALESCE(SUM(b{b}_clasa_2), 0) AS k2_{b}")
        sel.append(f"COALESCE(SUM(b{b}_clasa_4), 0) AS k4_{b}")
    sql = (f"SELECT {', '.join(sel)} FROM inregistrari_orare "
           f"WHERE an = ? AND luna = ? AND zi = ? GROUP BY contor")

    conn = get_traffic_db()._conn()
    rez = {}
    for r in conn.execute(sql, (data.year, data.month, data.day)).fetchall():
        n = int(r["n"] or 2)
        n = n if n in (2, 4, 6) else max(2, min(n, MAX_BENZI))
        rez[str(r["contor"])] = {
            "n":   n,
            "tot": {b: int(r[f"t{b}"]) for b in range(1, n + 1)},
            "c15": {b: int(r[f"c{b}"]) for b in range(1, n + 1)},
            "c2":  {b: int(r[f"k2_{b}"]) for b in range(1, n + 1)},
            "c4":  {b: int(r[f"k4_{b}"]) for b in range(1, n + 1)},
        }
    return rez


def _ultima_zi_cu_date(contori, data):
    """Pentru fiecare contor din `contori`, ultima zi (anterioara lui `data`)
    cu date inregistrate. Returneaza {contor: datetime.date}.

    Foloseste indexul cheii primare (contor, timestamp): MAX(timestamp) pe un
    contor se rezolva instant, fara sa citeasca tabelul."""
    conn = get_traffic_db()._conn()
    limita = data.strftime("%Y-%m-%d") + " 00:00:00"
    rez = {}
    for ct in contori:
        r = conn.execute(
            "SELECT MAX(timestamp) FROM inregistrari_orare "
            "WHERE contor = ? AND timestamp < ?", (ct, limita)).fetchone()
        if r and r[0]:
            rez[str(ct)] = datetime.datetime.strptime(r[0][:10], "%Y-%m-%d").date()
    return rez


# ── Construire rânduri raport ────────────────────────────────────────────────

def _construieste_randuri(data):
    contoare = get_contoare_db().get_all()      # doar contoarele ACTIVE
    zi = _citeste_zi(data)
    redenumite = _coduri_redenumite() if OMITE_REDENUMITE else set()
    fara_date = [str(c) for c in contoare
                 if str(c) not in zi and str(c).strip() not in redenumite]
    ultima_zi = _ultima_zi_cu_date(fara_date, data)
    try:
        verificari = citeste_verificari(data)
    except Exception:
        verificari = {}       # tabel/DB inaccesibil -> raportul se genereaza oricum

    randuri = []
    for ct_id, info in contoare.items():
        if str(ct_id).strip() in redenumite:
            continue
        d = zi.get(str(ct_id))
        if d is None:
            uz = ultima_zi.get(str(ct_id))
            piese_uz = (f"Ultima zi cu date inregistrate: {uz.strftime('%d.%m.%Y')}"
                        if uz else "Nu exista date inregistrate anterior")
            piese = _motiv_fara_date(verificari.get(str(ct_id)), info) + "\n" + piese_uz
            mod, probleme = "Fara date inregistrate in aceasta zi", "Fara date"
        else:
            mod_l, piese_l, are_problema = [], [], False
            incomplete = _benzi_incomplete(d["n"], d["tot"])
            for b in range(1, d["n"] + 1):
                st, extra = _stare_banda(d["tot"][b], d["c15"][b], d["c2"][b], d["c4"][b])
                inc = b in incomplete
                if inc and st == "T":            # totalizator + incomplet -> pe acelasi rand
                    mod_l.append(f"B{b} functioneaza totalizator si nu inregistreaza tot traficul")
                elif inc:                        # clasificator + incomplet -> doar mesajul de incomplet
                    mod_l.append(_MOD_INCOMPLET.format(b=b))
                else:
                    mod_l.append(_MOD_TXT[st].format(b=b))
                if st == "D":
                    piese_l.append(_PIESE_TXT["D"].format(b=b))
                for cheie in ("piezo", "bucla_iesire"):
                    if cheie in extra:
                        piese_l.append(_PIESE_TXT[cheie].format(b=b))
                if inc:
                    piese_l.append(_PIESE_INCOMPLET.format(b=b))
                if st != "C" or inc:
                    are_problema = True
            mod = "\n".join(mod_l)
            piese = "\n".join(piese_l)
            probleme = "Da" if are_problema else "Nu"

        randuri.append([
            info.get("DRDP", "") or "",
            info.get("Drum", "") or "",
            info.get("Pozitie_km", "") or "",
            ct_id,
            info.get("Localitate", "") or "",
            info.get("IP", "") or "",
            mod, piese, probleme,
        ])

    randuri.sort(key=lambda r: (str(r[0]), str(r[1]), str(r[3])))
    return randuri


# ── Scriere Excel ────────────────────────────────────────────────────────────

def _scrie_excel(randuri, data, cale):
    C_DARK, C_MID = "1F4E79", "2E75B6"
    F_DA, F_NU, F_NA = "FADBD8", "D5F5E3", "FFF2CC"
    thin = Side(style="thin", color="BFBFBF")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mod functionare"
    ws.sheet_view.showGridLines = False

    n_col = len(HEADERS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_col)
    t = ws.cell(1, 1, f"Modul de functionare al contoarelor active — {data.strftime('%d.%m.%Y')}")
    t.font = Font(name="Arial", size=13, bold=True, color="FFFFFF")
    t.fill = PatternFill("solid", start_color=C_DARK)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(2, c, h)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=C_MID)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTHS[c - 1]
    ws.row_dimensions[2].height = 30

    for r_idx, row in enumerate(randuri, 3):
        for c, val in enumerate(row, 1):
            cell = ws.cell(r_idx, c, val)
            cell.font = Font(name="Arial", size=10, bold=(c == 4))
            cell.border = brd
            left = c in (7, 8)
            cell.alignment = Alignment(horizontal="left" if left else "center",
                                       vertical="center", wrap_text=True)
        pc = ws.cell(r_idx, n_col)
        prob = row[-1]
        pc.fill = PatternFill("solid", start_color=(F_DA if prob == "Da" else
                                                    F_NU if prob == "Nu" else F_NA))
        pc.font = Font(name="Arial", size=10, bold=True,
                       color=("922B21" if prob == "Da" else
                              "1E8449" if prob == "Nu" else "7D6608"))
        n_lin = max(str(row[6]).count("\n"), str(row[7]).count("\n")) + 1
        ws.row_dimensions[r_idx].height = max(18, 14 * n_lin)

    ult = 2 + len(randuri)
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n_col)}{max(ult, 2)}"

    # Sumar
    nr_da = sum(1 for r in randuri if r[-1] == "Da")
    nr_nu = sum(1 for r in randuri if r[-1] == "Nu")
    nr_fd = sum(1 for r in randuri if r[-1] == "Fara date")
    s = ult + 2
    ws.cell(s, 1, f"Total contoare active: {len(randuri)}   |   "
                  f"Cu probleme: {nr_da}   |   Fara probleme: {nr_nu}   |   Fara date: {nr_fd}"
            ).font = Font(name="Arial", size=10, bold=True)

    wb.save(cale)


# ── API public ───────────────────────────────────────────────────────────────

def genereaza_raport(data, cale_iesire=None):
    """Generează raportul pentru `data` (datetime.date) și returnează calea fișierului."""
    randuri = _construieste_randuri(data)
    if cale_iesire is None:
        folder = os.path.join(CENTRAL_FILE_FOLDER, FOLDER_RAPOARTE)
        os.makedirs(folder, exist_ok=True)
        cale_iesire = os.path.join(
            folder, f"Raport_Mod_Functionare_{data.strftime('%Y-%m-%d')}.xlsx")
    _scrie_excel(randuri, data, cale_iesire)
    return cale_iesire


def _parse_data(s):
    s = s.strip()
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Data invalida: '{s}' (format asteptat: ZZ.LL.AAAA)")


def _radacina_tk():
    """Returneaza (root, creat_aici). Daca exista deja o aplicatie Tk (ex. PeekApp),
    o foloseste pe aceea si NU creeaza un al doilea tk.Tk()."""
    import tkinter as tk
    root = getattr(tk, "_default_root", None)
    if root is not None:
        return root, False
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root, True


def _cere_data_gui():
    """DOAR pentru rularea ca script (PyCharm). Din aplicatie se foloseste dialogul
    PeekApp._open_raport_dialog. Returneaza date sau None."""
    from tkinter import simpledialog, messagebox
    root, creat = _radacina_tk()
    implicit = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d.%m.%Y")
    try:
        while True:
            txt = simpledialog.askstring(
                "Raport mod functionare",
                "Data raportului (ZZ.LL.AAAA):",
                initialvalue=implicit, parent=root)
            if txt is None:            # Cancel
                return None
            try:
                return _parse_data(txt)
            except ValueError as e:
                messagebox.showerror("Data invalida", str(e), parent=root)
    finally:
        if creat:
            root.destroy()


def _anunta_gui(cale):
    from tkinter import messagebox
    root, creat = _radacina_tk()
    try:
        deschide = messagebox.askyesno(
            "Raport generat", f"Raport salvat:\n{cale}\n\nIl deschid acum?", parent=root)
    finally:
        if creat:
            root.destroy()
    if deschide:
        try:
            os.startfile(cale)         # Windows
        except AttributeError:
            pass


if __name__ == "__main__":
    if len(sys.argv) > 1:              # rulare din linie de comanda cu data
        zi = _parse_data(sys.argv[1])
        arg_out = sys.argv[2] if len(sys.argv) > 2 else None
        print(f"Raport generat: {genereaza_raport(zi, arg_out)}")
    else:                              # rulare din PyCharm: alegi data intr-o fereastra
        zi = _cere_data_gui()
        if zi is not None:
            cale = genereaza_raport(zi)
            print(f"Raport generat: {cale}")
            _anunta_gui(cale)
