# =============================================================================
# log_parser.py — Parsare fișiere .log VEK/ADR
# =============================================================================
# Exportă:
#   process_log_files(filepaths, output_dir, stop_event) → [{ path, id, ... }]
# =============================================================================

import os
import re
import pandas as pd

from config import LOG_CLASS_MAP, BAND_COLORS, MIN_ORE_ZI, CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER
from excel_report import (
    add_charts_and_formatting, precompute_localitati,
    read_mzl_manual_fallback_from_centralizator,
)
from database import get_traffic_db
from parallel_utils import (
    PoolHandle, ReportRunner, StopRequested, get_worker_count,
    submit_safe, get_safe, wait_one, timed_db, get_db_perf_breakdown,
)

# ══════════════════════════════════════════════════════════════════════════════
# PROCESARE FIȘIERE .LOG (VEK/ADR format)
# ══════════════════════════════════════════════════════════════════════════════
#
# Structură fișier .log:
#   linii cu câmpuri separate prin ";" după secțiunea "* Protocol:"
#   câmpuri: No; Det.Adr; Module; Direction; Class; Speed; Length;
#            Time gap; Busy time; Date; Time
#
# Clase vehicule → mapare pe Clasa_1..15 (identic cu fișierele .bin):
#   1=Motorbike, 2=Car, 3=Car with trailer, 4=Van, 5=Lorry,
#   6=Lorry with trailer, 7=Truck, 8=Bus, 15=Other
#
# Output:
#   - <Contor>_treceri_brute.csv  - toate înregistrările individuale
#   - Raport_Clase_Log_<Contor>.xlsx - Date Detaliate + toate analizele Peek
#   - Centralizator actualizat
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# PROCESARE FISIERE .LOG (VEK/ADR format)
# ══════════════════════════════════════════════════════════════════════════════
#
# Structura fisier .log:
#   linii cu campuri separate prin ";" dupa sectiunea "* Protocol:"
#   campuri: No; Det.Adr; Module; Direction; Class; Speed; Length;
#            Time gap; Busy time; Date; Time
#
# Banda = coloana Module (int): 1→B1, 2→B2, 3→B3, 4→B4, 5→B5, 6→B6
#
# Clase vehicule in Date Detaliate: B1_Motorbike, B1_Car, B1_Car_with_trailer,
#   B1_Van, B1_Lorry, B1_Lorry_with_trailer, B1_Truck, B1_Bus, B1_Other,
#   Total_B1, B2_Motorbike, ... Total_B2, etc.
#
# Mapare clasa text → index Clasa (identic cu .bin pentru centralizator):
#   Motorbike=1, Car=2, Car_with_trailer=3, Van=4, Lorry=5,
#   Lorry_with_trailer=6, Truck=7, Bus=8, Other=15
#
# Output:
#   - <Contor>_treceri_brute.csv  - toate inregistrarile individuale
#   - Raport_Clase_Log_<Contor>.xlsx - Date Detaliate (coloane cu nume) +
#     Media Zilnica Lunara + Media Zilnica Anuala
#   - Centralizator actualizat
# ══════════════════════════════════════════════════════════════════════════════

# Mapare clasa text VEK → index Clasa_1..15 (identic cu .bin)
LOG_CLASS_MAP = {
    "Motorbike":          1,
    "Car":                2,
    "Car with trailer":   3,
    "Van":                4,
    "Lorry":              5,
    "Lorry with trailer": 6,
    "Truck":              7,
    "Bus":                8,
    # 9..14 nu apar in VEK
    "Other":             15,
}


def _parse_log_file(filepath):
    """
    Parseaza un fisier .log VEK/ADR.
    Returneaza DataFrame brut (o linie per vehicul) sau DataFrame gol.
    Banda = int(Module): 1→B1, 2→B2, ..., 6→B6.
    site_id = tot ce e inainte de primul '_' din numele fisierului.
    """
    try:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    except Exception:
        return pd.DataFrame()

    if "* Protocol:" not in text:
        return pd.DataFrame()

    protocol_section = re.split(r"\* Protocol:", text, flags=re.IGNORECASE)[-1]
    lines = re.findall(r"^\s*\d+;.*", protocol_section, flags=re.MULTILINE)

    base    = os.path.splitext(os.path.basename(filepath))[0]
    site_id = base.split("_")[0]

    rows = []
    for line in lines:
        parts = [p.strip() for p in line.split(";")]
        parts = [p for p in parts if p]
        if len(parts) < 11:
            continue
        try:
            no, detadr, module, direction, cls, speed, length, \
                time_gap, busy_time, date_str, time_str = parts[:11]
            dt = pd.to_datetime(f"{date_str} {time_str}",
                                format="%d.%m.%Y %H:%M:%S")
            try:
                band = int(module)
                if band < 1 or band > 6:
                    band = 1
            except Exception:
                band = 1
            rows.append({
                "No":        int(no),
                "Det.Adr":   detadr,
                "Module":    band,
                "Direction": direction,
                "Class":     cls if cls in LOG_CLASS_MAP else "Other",
                "Speed":     float(speed),
                "Length":    float(length),
                "Time_gap":  float(time_gap),
                "Busy_time": float(busy_time),
                "Date":      date_str,
                "Time":      time_str,
                "Datetime":  dt,
                "SiteID":    site_id,
                "SourceFile":os.path.basename(filepath),
            })
        except Exception:
            continue

    return pd.DataFrame(rows)


def _build_log_hourly_df(df_raw, site_id):
    """
    Construieste DataFrame orar identic cu formatul .bin:
      Contor, Timestamp, Data_Ora, N_Benzi,
      B1_Clasa_1..15, Total_B1,
      B2_Clasa_1..15, Total_B2, ...,
      Total_General

    Banda = df_raw["Module"] (1..6).
    """
    if df_raw.empty:
        return pd.DataFrame()

    df = df_raw.copy()
    df["Clasa_idx"] = df["Class"].map(LOG_CLASS_MAP).fillna(15).astype(int)
    df["Hour"]      = df["Datetime"].dt.floor("h")

    # Numarul real de benzi din date
    n_bands = int(df["Module"].max()) if not df["Module"].empty else 1
    n_bands = max(1, min(n_bands, 6))

    rows = []
    for hour_ts, grp in df.groupby("Hour"):
        row = {
            "Contor":    site_id,
            "Timestamp": hour_ts,
            "Data_Ora":  hour_ts.strftime("%d.%m.%Y %H:%M"),
            "N_Benzi":   n_bands,
        }
        total_general = 0
        for b in range(1, n_bands + 1):
            grp_b  = grp[grp["Module"] == b]
            counts = grp_b["Clasa_idx"].value_counts()
            total_b = 0
            for cls_idx in list(range(1, 9)) + [15]:
                val = int(counts.get(cls_idx, 0))
                row[f"B{b}_Clasa_{cls_idx}"] = val
                total_b += val
            row[f"Total_B{b}"] = total_b
            total_general += total_b
        row["Total_General"] = total_general
        rows.append(row)

    return pd.DataFrame(rows)


def _prepare_log_counter(site_id, files):
    """
    Pasul CPU-intens pentru UN contor (rulează într-un proces worker):
    parsare .log → CSV brut → DataFrame orar → sortare/deduplicare → curățare benzi.
    Întoarce dict sau None (fără date valide).
    """
    # ── 1. Parsare ────────────────────────────────────────────────────────
    frames = []
    for fp in files:
        df_raw = _parse_log_file(fp)
        if not df_raw.empty:
            frames.append(df_raw)

    if not frames:
        return None

    df_all = pd.concat(frames, ignore_index=True)
    df_all = df_all.sort_values("Datetime").reset_index(drop=True)

    # ── Calculăm perioada per fișier sursă ───────────────────────────────
    file_periods = {}
    for fp in files:
        fname = os.path.basename(fp)
        df_f = df_all[df_all["SourceFile"] == fname]
        if not df_f.empty:
            p_min = df_f["Datetime"].min().strftime("%d.%m.%Y")
            p_max = df_f["Datetime"].max().strftime("%d.%m.%Y")
            file_periods[fname] = (p_min, p_max)

    # ── 2. CSV brut — salvat în folderul SURSĂ al fișierelor .log ────────
    # Folderul sursă = folderul comun al fișierelor pentru acest site_id.
    # Dacă fișierele sunt în foldere diferite, folosim folderul primului fișier.
    sursa_folders = list(dict.fromkeys(os.path.dirname(os.path.abspath(fp)) for fp in files))
    csv_folder = sursa_folders[0]  # folderul rădăcină al primului fișier sursă
    csv_path = os.path.join(csv_folder, f"{site_id}_treceri_brute.csv")
    df_all.to_csv(csv_path, index=False, encoding="utf-8-sig")

    # ── 3. DataFrame orar (format identic .bin) ───────────────────────────
    df_hourly = _build_log_hourly_df(df_all, site_id)
    if df_hourly.empty:
        return None

    n_lanes = int(df_all["Module"].max()) if not df_all["Module"].empty else 1
    n_lanes = max(1, min(n_lanes, 6))

    # Sortare + deduplicare (identic cu .bin)
    df_hourly["_dt"] = pd.to_datetime(
        df_hourly["Data_Ora"], format="%d.%m.%Y %H:%M", errors="coerce")
    df_hourly = df_hourly.sort_values(["Contor", "_dt"])
    df_hourly = df_hourly.drop_duplicates(subset=["Contor", "_dt"], keep="last")
    df_hourly = df_hourly.drop(columns=["_dt"], errors="ignore")
    df_hourly = df_hourly.reset_index(drop=True)

    # Eliminam benzile goale (identic cu .bin)
    cols_to_drop = []
    for b in range(n_lanes + 1, 7):
        for c in range(1, 16):
            col = f"B{b}_Clasa_{c}"
            if col in df_hourly.columns:
                cols_to_drop.append(col)
        if f"Total_B{b}" in df_hourly.columns:
            cols_to_drop.append(f"Total_B{b}")
    if cols_to_drop:
        df_hourly = df_hourly.drop(columns=cols_to_drop)

    # Recalculam Total_General
    band_tot_cols = [f"Total_B{b}" for b in range(1, n_lanes + 1)
                     if f"Total_B{b}" in df_hourly.columns]
    df_hourly["Total_General"] = df_hourly[band_tot_cols].sum(axis=1)

    return {"df_hourly": df_hourly, "n_lanes": n_lanes,
            "file_periods": file_periods, "csv_path": csv_path,
            "band_tot_cols": band_tot_cols}


def process_log_files(filepaths, output_dir=None, stop_event=None,
                      progress_callback=None, warning_callback=None):
    """
    Proceseaza fisiere .log VEK/ADR.
    progress_callback(site_id, n_ore, idx, total_contoare) - apelat dupa
    fiecare contor procesat si scris in SQLite + Excel (in ordinea terminarii).
    warning_callback(msg) - optional, pentru erori la un contor.

    Paralelism (mai multe nuclee):
      * parsarea + constructia DataFrame-ului orar, per contor, in procese worker;
      * scrierea in SQLite si citirea istoricului, in procesul principal
        (un singur scriitor in DB), pe masura ce contoarele sunt gata;
      * generarea Excel + grafice, in procese worker.
    Numarul de procese: vezi parallel_utils.get_worker_count().
    """
    if not filepaths:
        return None

    # output_dir = folderul pentru Raport_Clase_VEK (vine din app.py)
    # treceri_brute se salvează separat, în folderul sursă al fiecărui contor
    if output_dir is None:
        output_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)
    os.makedirs(output_dir, exist_ok=True)

    # Grupam fisierele pe site_id (logica VEK: tot ce e inainte de primul "_")
    contoare_files = {}
    for fp in filepaths:
        base = os.path.splitext(os.path.basename(fp))[0]
        sid  = base.split("_")[0]
        contoare_files.setdefault(sid, []).append(fp)

    _total_ct = len(contoare_files)
    pool = PoolHandle(get_worker_count(_total_ct) if _total_ct >= 2 else 1)
    runner = None

    try:
        try:
            tdb = get_traffic_db()
        except Exception as _e_tdb:
            tdb = None
            print(f"  [WARN] SQLite indisponibil: {_e_tdb}")

        localitati = precompute_localitati(list(contoare_files.keys()))
        # Orele scrise ACUM în SQLite per contor; se transmit callback-ului de
        # progres ca al 5-lea argument DOAR dacă acesta îl acceptă.
        _scrise_now = {}
        _cb = progress_callback
        if _cb is not None:
            import inspect
            try:
                _params = list(inspect.signature(_cb).parameters.values())
                _accepts5 = (len(_params) >= 5 or
                             any(p.kind == p.VAR_POSITIONAL for p in _params))
            except (TypeError, ValueError):
                _accepts5 = False

            def _progress_wrap(sid, n_ore, idx, total):
                if _accepts5:
                    return _cb(sid, n_ore, idx, total, _scrise_now.get(sid))
                return _cb(sid, n_ore, idx, total)
            progress_callback = _progress_wrap
        runner = ReportRunner(pool, stop_event, progress_callback, _total_ct,
                              warn=warning_callback)

        # Pregatirea tuturor contoarelor porneste imediat (in paralel);
        # le consumam in ordine, ca scrierea in DB sa se suprapuna cu
        # munca proceselor worker.
        prep = []
        for site_id, files in contoare_files.items():
            if stop_event and stop_event.is_set():
                runner.abort()
                return None
            prep.append((site_id, files, submit_safe(pool, _prepare_log_counter, site_id, files)))

        for _idx_ct, (site_id, files, fut) in enumerate(prep, 1):
            wait_one(fut, stop_event)                    # ridica StopRequested la anulare
            res = get_safe(fut, _prepare_log_counter, site_id, files)
            if res.get("log"):
                print(res["log"], end="" if res["log"].endswith("\n") else "\n")
            if not res["ok"]:
                msg = (res.get("error") or "").strip().splitlines()
                msg = msg[-1] if msg else "necunoscuta"
                (warning_callback or print)(f"  ✗ [{site_id}] eroare la citirea fisierelor .log: {msg}")
                print(res.get("error", ""))
                continue
            prepared = res["result"]
            if prepared is None:
                continue

            df_hourly     = prepared["df_hourly"]
            n_lanes       = prepared["n_lanes"]
            file_periods  = prepared["file_periods"]
            csv_path      = prepared["csv_path"]
            band_tot_cols = prepared["band_tot_cols"]

            # ── 3. Scriere în SQLite (date orare agregate) ────────────────
            try:
                with timed_db("scriere"):
                    n_scrise = tdb.upsert_hourly_df(
                        df_hourly, site_id,
                        tip_sursa="VEK",
                        source_files=files,
                        source_file_periods=file_periods
                    )
                _scrise_now[site_id] = n_scrise
                print(f"  💾 SQLite ← [{site_id}]  {n_scrise} rânduri orare (VEK)")
            except Exception as _e_db:
                print(f"  [WARN] SQLite write eroare [{site_id}]: {_e_db}")

            # ── 4. Citire din SQLite — toate datele acumulate pentru site_id ─
            # Identic cu .bin: Excel-ul reflectă ÎNTREAGA perioadă din DB,
            # nu doar fișierele procesate acum.
            df_din_db = None
            all_source_files_db = list(files)  # fallback: fișierele curente
            try:
                if tdb is not None:
                    with timed_db("citire"):
                        df_din_db = tdb.get_hourly_df(site_id)
            except Exception as _e_read:
                print(f"  [WARN] Citire DB eroare [{site_id}]: {_e_read}")

            if df_din_db is not None and not df_din_db.empty:
                # Folosim datele complete din DB
                df_for_excel = df_din_db.copy()

                # ── Detectăm numărul real de benzi din coloana N_Benzi din DB ──
                # Schema SQLite stochează B1..B6 hardcodat — benzile dincolo de
                # N_Benzi sunt artefacte de schemă (zerouri), nu benzi reale.
                # Nu folosim sum()==0 ca criteriu, pentru că o bandă defectă
                # (zero vehicule pe toată perioada) e tot o bandă reală a contorului.
                n_lanes_db = n_lanes  # fallback: valoarea din parsarea curentă
                if "N_Benzi" in df_for_excel.columns:
                    n_benzi_vals = df_for_excel["N_Benzi"].dropna()
                    if not n_benzi_vals.empty:
                        n_lanes_db = int(n_benzi_vals.max())

                # Eliminăm coloanele benzilor dincolo de n_lanes_db (artefacte schemă)
                cols_schema_drop = []
                for b in range(n_lanes_db + 1, 7):
                    tcol = f"Total_B{b}"
                    if tcol in df_for_excel.columns:
                        cols_schema_drop.append(tcol)
                    for cls in list(range(1, 9)) + [15]:
                        ccol = f"B{b}_Clasa_{cls}"
                        if ccol in df_for_excel.columns:
                            cols_schema_drop.append(ccol)
                if cols_schema_drop:
                    df_for_excel = df_for_excel.drop(columns=cols_schema_drop, errors="ignore")

                n_lanes = n_lanes_db
                band_tot_cols = sorted(
                    [c for c in df_for_excel.columns
                     if c.startswith("Total_B") and c != "Total_General"],
                    key=lambda x: int(x.replace("Total_B", ""))
                )
                if band_tot_cols:
                    df_for_excel["Total_General"] = df_for_excel[band_tot_cols].sum(axis=1)

                # Colectăm lista tuturor fișierelor sursă înregistrate în DB
                try:
                    with timed_db("surse"):
                        all_source_files_db = tdb.get_source_files(site_id) or list(files)
                except Exception:
                    all_source_files_db = list(files)
            else:
                # Fallback: folosim datele tocmai parsate (prima rulare sau DB indisponibil)
                df_for_excel = df_hourly.copy()

            # ── 5. Excel: trimis către un proces worker ───────────────────
            output_fn = os.path.join(output_dir, f"Raport_Clase_VEK_{site_id}.xlsx")

            all_band_totals = sum(df_for_excel[c].sum() for c in band_tot_cols
                                  if c in df_for_excel.columns)

            # Implicit: goale, nu None — un worker primind None ar încerca să
            # citească el însuși din rețea (vezi read_mzl_manual_fallback_...).
            # Aici, în procesul principal, un eșec are un fallback sigur:
            # foaia "Prelucrare manuala" a centralizatorului, citită O SINGURĂ
            # DATĂ per contor, niciodată din interiorul unui worker.
            _mzl_manual = {}
            _live_src   = []
            if tdb is not None:
                try:
                    with timed_db("mzl_manual"):
                        _mzl_manual = tdb.get_mzl_manual(str(site_id)) or {}
                except Exception as _e_mzl:
                    print(f"  [WARN] Citire MZL manual eroare [{site_id}]: {_e_mzl}")
                    _mzl_manual = read_mzl_manual_fallback_from_centralizator(site_id)
                try:
                    with timed_db("surse"):
                        _live_src = tdb.get_source_files_periods(str(site_id)) or []
                except Exception as _e_live:
                    print(f"  [WARN] Citire surse fisier eroare [{site_id}]: {_e_live}")
                    _live_src = []
            else:
                _mzl_manual = read_mzl_manual_fallback_from_centralizator(site_id)

            suma_b1 = df_for_excel["Total_B1"].sum() if "Total_B1" in df_for_excel.columns else 0
            suma_b2 = df_for_excel["Total_B2"].sum() if "Total_B2" in df_for_excel.columns else 0
            n_randuri = len(df_for_excel)

            job = {
                "site_id": site_id,
                "output_fn": output_fn,
                "tmp_fn": os.path.join(output_dir, f".tmp_{os.path.basename(output_fn)}"),
                "df": df_for_excel,
                "has_traffic": all_band_totals > 0,
                "zero_msg": f"Atentie: Contorul LOG {site_id} are trafic zero.",
                "source_files": all_source_files_db,
                "source_file_periods": file_periods,
                "mzl_manual": _mzl_manual,
                "live_sources": _live_src,
                "localitate_site": localitati.get(str(site_id)),
            }
            rezultat = {
                "path":       output_fn,
                "id":         site_id,
                "randuri":    n_randuri,
                "b1":         int(suma_b1),
                "b2":         int(suma_b2),
                "n_lanes":    n_lanes,
                "source_files": all_source_files_db,
                "csv_path":   csv_path,
            }
            runner.submit(job, order=_idx_ct, rezultat=rezultat,
                          site_id=site_id, randuri=n_randuri)

        rezultate = runner.finish()
        if rezultate is None:          # anulat de utilizator
            return None

    except StopRequested:
        (runner or pool).abort()
        return None
    except BaseException:
        (runner or pool).abort()
        raise

    return rezultate if rezultate else None



