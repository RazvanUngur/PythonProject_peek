# =============================================================================
# log_parser.py — Parsare fișiere .log VEK/ADR
# =============================================================================
# Exportă:
#   process_log_files(filepaths, output_dir, stop_event) → [{ path, id, ... }]
# =============================================================================

import os
import re
import pandas as pd

from config import LOG_CLASS_MAP, BAND_COLORS, MIN_ORE_ZI
from excel_report import add_charts_and_formatting
from database import get_traffic_db

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


def process_log_files(filepaths, output_dir=None, stop_event=None,
                      progress_callback=None):
    """
    Proceseaza fisiere .log VEK/ADR.
    progress_callback(site_id, n_ore, idx, total_contoare) - apelat dupa
    fiecare contor procesat si scris in SQLite + Excel.
    """
    if not filepaths:
        return None

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(filepaths[0]))
    os.makedirs(output_dir, exist_ok=True)

    # Grupam fisierele pe site_id (logica VEK: tot ce e inainte de primul "_")
    contoare_files = {}
    for fp in filepaths:
        base = os.path.splitext(os.path.basename(fp))[0]
        sid  = base.split("_")[0]
        contoare_files.setdefault(sid, []).append(fp)

    rezultate   = []
    _total_ct   = len(contoare_files)

    for _idx_ct, (site_id, files) in enumerate(contoare_files.items(), 1):
        if stop_event and stop_event.is_set():
            return None

        # ── 1. Parsare ────────────────────────────────────────────────────────
        frames = []
        for fp in files:
            if stop_event and stop_event.is_set():
                return None
            df_raw = _parse_log_file(fp)
            if not df_raw.empty:
                frames.append(df_raw)

        if not frames:
            continue

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

        # ── 2. CSV brut ───────────────────────────────────────────────────────
        csv_path = os.path.join(output_dir, f"{site_id}_treceri_brute.csv")
        df_all.to_csv(csv_path, index=False, encoding="utf-8-sig")

        # ── 3. DataFrame orar (format identic .bin) ───────────────────────────
        df_hourly = _build_log_hourly_df(df_all, site_id)
        if df_hourly.empty:
            continue

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

        # ── 3. Scriere în SQLite (date orare agregate) ────────────────────────
        try:
            tdb = get_traffic_db()
            n_scrise = tdb.upsert_hourly_df(
                df_hourly, site_id,
                tip_sursa="VEK",
                source_files=files
            )
            print(f"  💾 SQLite ← [{site_id}]  {n_scrise} rânduri orare (VEK)")
        except Exception as _e_db:
            print(f"  [WARN] SQLite write eroare [{site_id}]: {_e_db}")

        # ── 4. Excel cu toate sheeturile ──────────────────────────────────────
        output_fn = os.path.join(output_dir, f"Raport_Clase_VEK_{site_id}.xlsx")

        df_export = df_hourly.drop(columns=["N_Benzi"], errors="ignore")
        df_export = df_export.reset_index(drop=True)

        df_export.to_excel(output_fn, index=False, sheet_name="Date Detaliate")

        all_band_totals = sum(df_export[c].sum() for c in band_tot_cols
                              if c in df_export.columns)

        df_for_charts = df_hourly.copy()
        df_for_charts["Timestamp"] = pd.to_datetime(
            df_for_charts["Data_Ora"], format="%d.%m.%Y %H:%M", errors="coerce")

        if all_band_totals > 0:
            add_charts_and_formatting(output_fn, df_for_charts, site_id,
                                      source_files=files,
                                      source_file_periods=file_periods)
        else:
            print(f"Atentie: Contorul LOG {site_id} are trafic zero.")

        suma_b1 = df_hourly["Total_B1"].sum() if "Total_B1" in df_hourly.columns else 0
        suma_b2 = df_hourly["Total_B2"].sum() if "Total_B2" in df_hourly.columns else 0

        rezultate.append({
            "path":    output_fn,
            "id":      site_id,
            "randuri": len(df_export),
            "b1":      int(suma_b1),
            "b2":      int(suma_b2),
            "n_lanes": n_lanes,
            "source_files": files,
        })

        # Notifică GUI progresul după fiecare contor complet
        if progress_callback:
            try:
                progress_callback(site_id, len(df_export), _idx_ct, _total_ct)
            except Exception:
                pass

    return rezultate if rezultate else None



