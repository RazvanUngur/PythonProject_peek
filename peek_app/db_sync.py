# =============================================================================
# db_sync.py — sincronizare "copiază local / lucrează / trimite înapoi"
# =============================================================================
# Motivație: trafic.db locuiește pe un server de rețea, accesat de mai mulți
# colegi (câte unul pe rând, impus de lacătul existent .procesare.lock din
# app.py). SQLite în modul WAL peste rețea (SMB) e foarte lent — fiecare
# citire/scriere așteaptă rețeaua, ba chiar a produs blocaje de acces pentru
# alți utilizatori. Măsurători: cu baza pe rețea, 3 contoare = ~114s doar
# SQLite; cu baza locală, ~4s.
#
# Soluția: cât timp lacătul e al nostru (deci suntem singurul utilizator activ),
# copiem baza de date local, lucrăm pe copie (rapid), apoi copiem rezultatul
# înapoi pe server înainte să eliberăm lacătul. La fel pentru rapoartele Excel
# generate — scrise întâi local, copiate pe server abia la final.
#
# contoare.db NU este afectată — rămâne pe rețea ca înainte (o ating doar
# centralizator.py, nu fluxul de procesare .bin/.log).
#
# Utilizare tipică (din app.py, în jurul unei sesiuni de procesare):
#
#     import db_sync
#     try:
#         db_sync.checkout_database(log=self._log)
#         local_reports_dir = db_sync.prepare_local_reports_dir()
#         sync_activ = True
#     except Exception as e:
#         self._log(f"[WARN] sincronizare locală indisponibilă: {e}")
#         local_reports_dir = <folderul de rețea, ca înainte>
#         sync_activ = False
#     ...
#     # foloseşte local_reports_dir ca output_dir pentru process_multiple_files/
#     # process_log_files ...
#     ...
#     if sync_activ:
#         db_sync.sync_reports_to_server(local_reports_dir, server_reports_dir, log=self._log)
#         db_sync.commit_database(log=self._log)
# =============================================================================

import os
import shutil
import tempfile
import time

import database
from config import CENTRAL_FILE_FOLDER

try:
    import config as _cfg
except Exception:            # pragma: no cover
    _cfg = None


def _default_log(msg):
    print(msg)


# ─────────────────────────────────────────────────────────────────────────────
# Pragul (număr de contoare) sub care NU merită copierea completă a bazei.
# O copiere completă are un cost FIX (proporțional cu mărimea întregii baze,
# nu cu ce se procesează acum), în timp ce lucrul direct pe rețea are un cost
# aproximativ proporțional cu numărul de contoare atinse. Sub prag, e mai
# ieftin să lucrăm direct pe rețea; peste, copierea completă câștigă.
# Baza va tot crește — dacă pragul nu mai pare potrivit, suprascrie-l cu
# CESTRIN_LOCAL_COPY_THRESHOLD=N (variabilă de mediu) sau
# LOCAL_COPY_THRESHOLD_CONTOARE = N în config.py, fără să modifici codul.
# Implicit 15, calculat din: copiere completă ~611s (5.2 GB) vs ~38s/contor
# direct pe rețea → pragul de rentabilitate e ~16 contoare.
# ─────────────────────────────────────────────────────────────────────────────
def _get_local_copy_threshold():
    v = (os.environ.get("CESTRIN_LOCAL_COPY_THRESHOLD")
         or getattr(_cfg, "LOCAL_COPY_THRESHOLD_CONTOARE", None))
    try:
        return int(v) if v else 15
    except (TypeError, ValueError):
        return 15


LOCAL_COPY_THRESHOLD = _get_local_copy_threshold()


def get_local_cache_dir():
    """Folder local, per-mașină, pentru copia de lucru a bazei și rapoartelor."""
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    path = os.path.join(base, "CESTRIN_TrafficApp_cache")
    os.makedirs(path, exist_ok=True)
    return path


def _local_db_folder():
    path = os.path.join(get_local_cache_dir(), "sqlite")
    os.makedirs(path, exist_ok=True)
    return path


def local_reports_folder():
    return os.path.join(get_local_cache_dir(), "Rapoarte")


# ─────────────────────────────────────────────────────────────────────────────
# Baza de date (trafic.db)
# ─────────────────────────────────────────────────────────────────────────────
def _is_lock_err(e):
    """Blocare temporară de fișier (Windows): PermissionError / WinError 32,33."""
    return isinstance(e, PermissionError) or getattr(e, "winerror", None) in (32, 33)


def _retry_io(fn, what, log, max_wait=180):
    """
    Rulează fn(); dacă Windows raportează fișierul ca blocat de alt proces
    (SQLite blochează scurt/ lung zone din trafic.db — ex. harta, un coleg sau
    o instanță veche a aplicației), reîncearcă cu pauze crescătoare, până la
    max_wait secunde. Altă eroare => ridicată imediat.
    """
    t0 = time.perf_counter()
    delay, warned = 0.2, False
    while True:
        try:
            return fn()
        except OSError as e:
            if not _is_lock_err(e):
                raise
            if time.perf_counter() - t0 > max_wait:
                raise OSError(
                    f"{what}: fișierul rămâne blocat de alt proces de peste "
                    f"{max_wait}s ({e}). Închide harta din browser, alte "
                    f"instanțe ale aplicației și procesele python rămase, "
                    f"apoi reîncearcă.") from e
            if not warned:
                log(f"      ⏳ {what}: fișier blocat temporar de alt proces — reîncerc...")
                warned = True
            time.sleep(delay)
            delay = min(delay * 1.5, 2.0)


def _copy_with_progress(src, dst, log, report_every=3.0):
    """Copiază src -> dst pe bucăți, afișând progres periodic (implicit la
    fiecare ~3s). Fiecare citire/scriere se reîncearcă la blocări temporare
    (SQLite ține lacăte pe zone din fișier, iar la ~1 GB există un octet de
    lacăt „PENDING” pe care orice conexiune deschisă îl blochează scurt —
    exact acolo cădea copia cu «Permission denied» / WinError 33).
    Păstrează metadatele ca shutil.copy2."""
    total = os.path.getsize(src)
    total_mb = total / 1_048_576
    if total_mb < 5:
        # Fișier mic — o copiere simplă e instantanee, fără progres.
        _retry_io(lambda: shutil.copy2(src, dst), os.path.basename(src), log)
        return
    log(f"    {os.path.basename(src)}: {total_mb:.0f} MB...")
    t0 = time.perf_counter()
    last_report = t0
    copied = 0
    CHUNK = 16 * 1024 * 1024
    name = os.path.basename(src)
    with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
        while True:
            def _read(off=copied):
                fsrc.seek(off)
                return fsrc.read(CHUNK)
            buf = _retry_io(_read, f"citire {name}", log)
            if not buf:
                break

            def _write(off=copied, data=buf):
                fdst.seek(off)
                fdst.write(data)
            _retry_io(_write, f"scriere locală {name}", log)
            copied += len(buf)
            now = time.perf_counter()
            if now - last_report >= report_every:
                pct = 100 * copied / total if total else 100
                mb_s = (copied / 1_048_576) / max(now - t0, 0.001)
                log(f"      ... {copied / 1_048_576:.0f}/{total_mb:.0f} MB "
                    f"({pct:.0f}%, {mb_s:.1f} MB/s)")
                last_report = now
    _retry_io(lambda: shutil.copystat(src, dst), f"metadate {name}", log)
    dt = time.perf_counter() - t0
    mb_s = total_mb / max(dt, 0.001)
    log(f"    ✓ copiat în {dt:.1f}s ({mb_s:.1f} MB/s).")


def checkout_database(log=None):
    """
    Copiază trafic.db (+ -wal/-shm dacă există) de pe server local, apoi
    redirecționează get_traffic_db() către copia locală.
    Ridică o excepție dacă serverul nu e accesibil sau copierea eșuează —
    apelantul decide fallback-ul (de regulă: lucrează direct pe rețea, ca
    înainte de sincronizare).
    """
    log = log or _default_log
    server_db = database.TRAFFIC_DB   # <CENTRAL_FILE_FOLDER>/<SQLITE_FOLDER>/trafic.db

    local_folder = _local_db_folder()
    # Curățăm orice rămășiță dintr-o sesiune anterioară neterminată corect,
    # ca să nu amestecăm date vechi cu cele proaspăt copiate.
    for suffix in ("", "-wal", "-shm"):
        stale = os.path.join(local_folder, "trafic.db" + suffix)
        if os.path.exists(stale):
            try:
                os.remove(stale)
            except Exception as _e_rm:
                raise OSError(
                    f"Nu pot șterge copia locală veche {stale} ({_e_rm}). "
                    f"Probabil un proces python rămas deschis o folosește — "
                    f"închide-l din Task Manager și reîncearcă.") from _e_rm

    if not os.path.exists(server_db):
        # Bază inexistentă încă pe server (prima rulare) — TrafficDB() o va
        # crea la prima conectare, local; e în regulă, nu e o eroare.
        log("  📥 Nicio bază existentă pe server încă — pornesc una nouă local.")
    else:
        log("  📥 Copiez baza de date pe server local — poate dura, în "
            "funcție de mărime și viteza rețelei...")
        copied = []
        # -shm NU se copiază: e doar un index în memorie, SQLite îl recreează
        # singur la prima deschidere (și e fișierul cel mai des blocat).
        for suffix in ("", "-wal"):
            src = server_db + suffix
            if not os.path.exists(src):
                continue
            dst = os.path.join(local_folder, os.path.basename(src))
            for attempt in (1, 2):
                size_before = os.path.getsize(src)
                _copy_with_progress(src, dst, log)
                if os.path.getsize(src) == size_before:
                    break
                if attempt == 1:
                    log("  ⚠ Baza s-a modificat pe server în timpul copierii — "
                        "o copiez din nou.")
                else:
                    raise OSError("trafic.db se modifică pe server în timpul "
                                  "copierii (alt proces scrie în el).")
            copied.append(os.path.basename(src))
        log(f"  📥 Baza de date copiată local ({', '.join(copied)}).")

    local_db_path = os.path.join(local_folder, "trafic.db")
    if os.path.exists(local_db_path):
        # Test rapid de sănătate a copiei (citește schema; nu parcurge toată baza).
        import sqlite3
        try:
            _c = sqlite3.connect(local_db_path, timeout=30)
            try:
                _c.execute("SELECT count(*) FROM sqlite_master").fetchone()
            finally:
                _c.close()
        except Exception as _e_chk:
            raise OSError(f"Copia locală a bazei nu se poate deschide: {_e_chk}") from _e_chk
    database.set_traffic_db_path(local_db_path)
    return local_db_path


def commit_database(log=None):
    """
    Face checkpoint (TRUNCATE — fuzionează -wal în fișierul principal, ca
    fișierul copiat pe server să fie curat, fără -wal/-shm rămase), închide
    conexiunea locală, copiază trafic.db local înapoi pe server (scriere
    într-un fișier temporar + înlocuire atomică), apoi revine la calea de
    rețea implicită.
    NU șterge copia locală — rămâne ca sursă de recuperare dacă ceva
    întrerupe sincronizarea următoare.
    """
    log = log or _default_log
    tdb = database.get_traffic_db()   # instanța curentă — trebuie să fie cea locală
    try:
        conn = tdb._conn()
        conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        conn.commit()
    except Exception as e:
        log(f"  [WARN] wal_checkpoint eșuat ({e}) — continui oricum, "
            f"copiez fișierele existente.")
    local_db_path = tdb.db_path
    tdb.close()

    server_db = database.TRAFFIC_DB
    os.makedirs(os.path.dirname(server_db), exist_ok=True)
    log("  📤 Copiez baza de date înapoi pe server...")
    tmp_dst = server_db + ".tmp_sync"
    try:
        _copy_with_progress(local_db_path, tmp_dst, log)
        _retry_io(lambda: os.replace(tmp_dst, server_db),
                  "înlocuire trafic.db pe server", log, max_wait=120)
    finally:
        if os.path.exists(tmp_dst):
            try:
                os.remove(tmp_dst)
            except Exception:
                pass
    # Curățăm eventuale -wal/-shm vechi rămase pe server (nu mai au sens
    # lângă un fișier proaspăt copiat, curat, fără jurnal neaplicat).
    for suffix in ("-wal", "-shm"):
        p = server_db + suffix
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
    log("  📤 Baza de date copiată înapoi pe server.")
    database.clear_traffic_db_path()


# ─────────────────────────────────────────────────────────────────────────────
# Rapoarte Excel
# ─────────────────────────────────────────────────────────────────────────────
def prepare_local_reports_dir():
    """Folder local, gol, pentru rapoartele generate în această sesiune."""
    path = local_reports_folder()
    if os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)
    os.makedirs(path, exist_ok=True)
    return path


def sync_reports_to_server(local_dir, server_dir, log=None):
    """
    Copiază toate fișierele din local_dir în server_dir (suprascrie).
    O eroare la un fișier nu oprește restul. Întoarce lista fișierelor care
    NU s-au putut copia — apelantul trebuie să avertizeze clar utilizatorul
    dacă lista nu e goală (înseamnă că acel raport a rămas doar local).
    """
    log = log or _default_log
    os.makedirs(server_dir, exist_ok=True)
    failed = []
    if not os.path.isdir(local_dir):
        return failed
    fisiere = sorted(f for f in os.listdir(local_dir)
                     if os.path.isfile(os.path.join(local_dir, f)))
    for fname in fisiere:
        src = os.path.join(local_dir, fname)
        dst = os.path.join(server_dir, fname)
        try:
            tmp_dst = dst + ".tmp_sync"
            shutil.copy2(src, tmp_dst)
            os.replace(tmp_dst, dst)
        except Exception as e:
            failed.append((fname, str(e)))
    ok = len(fisiere) - len(failed)
    if fisiere:
        log(f"  📤 Rapoarte copiate pe server: {ok}/{len(fisiere)}.")
    return failed
