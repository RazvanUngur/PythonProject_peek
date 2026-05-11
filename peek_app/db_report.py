# =============================================================================
# db_report.py — Generare rapoarte Excel direct din SQLite
# =============================================================================
# Exportă:
#   generate_report_from_db(contor, an, output_dir) → path Excel
#   generate_all_reports_from_db(output_dir, an)    → list rezultate
# =============================================================================

import os
import pandas as pd

from database import get_traffic_db, get_contoare_db
from excel_report import add_charts_and_formatting
from centralizator import update_centralizator
from config import CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME


def generate_report_from_db(contor: str, an: int = None,
                             output_dir: str = None,
                             tip_sursa: str = None) -> str | None:
    """
    Generează Raport_Clase_PEEK/VEK_<contor>.xlsx dintr-o interogare SQLite.
    Returnează calea fișierului Excel generat, sau None dacă nu există date.
    """
    if output_dir is None:
        output_dir = CENTRAL_FILE_FOLDER
    os.makedirs(output_dir, exist_ok=True)

    tdb = get_traffic_db()
    df  = tdb.get_hourly_df(contor, an=an)

    if df is None or df.empty:
        print(f"[RAPORT] Nicio dată în DB pentru contorul {contor}"
              + (f" / {an}" if an else ""))
        return None

    if tip_sursa is None:
        row = tdb._conn().execute(
            "SELECT tip_sursa FROM inregistrari_orare "
            "WHERE contor = ? LIMIT 1", (contor,)
        ).fetchone()
        tip_sursa = row["tip_sursa"] if row else "PEEK"

    prefix    = "Peek" if tip_sursa == "PEEK" else "VEK"
    output_fn = os.path.join(output_dir, f"Raport_Clase_{prefix}_{contor}.xlsx")

    n_benzi = int(df["N_Benzi"].iloc[0]) if "N_Benzi" in df.columns else 2
    cols_to_drop = []
    for b in range(n_benzi + 1, 7):
        for c in list(range(1, 9)) + [15]:
            col = f"B{b}_Clasa_{c}"
            if col in df.columns:
                cols_to_drop.append(col)
        if f"Total_B{b}" in df.columns:
            cols_to_drop.append(f"Total_B{b}")
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop)

    band_tot_cols = [f"Total_B{b}" for b in range(1, n_benzi + 1)
                     if f"Total_B{b}" in df.columns]
    df["Total_General"] = df[band_tot_cols].sum(axis=1)

    df_export = df.drop(columns=["N_Benzi"], errors="ignore").reset_index(drop=True)
    df_export.to_excel(output_fn, index=False, sheet_name="Date Detaliate")

    all_band_totals = sum(df_export[c].sum() for c in band_tot_cols
                          if c in df_export.columns)
    if all_band_totals > 0:
        df_for_charts = df.copy()
        if "Timestamp" not in df_for_charts.columns:
            df_for_charts["Timestamp"] = pd.to_datetime(
                df_for_charts["Data_Ora"], format="%d.%m.%Y %H:%M", errors="coerce")
        add_charts_and_formatting(output_fn, df_for_charts, contor)
    else:
        print(f"[RAPORT] Contorul {contor} are trafic zero.")

    return output_fn


def generate_all_reports_from_db(output_dir: str = None,
                                  an: int = None) -> list:
    """Generează rapoarte Excel pentru toate contoarele din DB."""
    tdb      = get_traffic_db()
    contoare = tdb.get_contoare_disponibile()
    rezultate = []

    for contor in contoare:
        path = generate_report_from_db(contor, an=an, output_dir=output_dir)
        if path:
            df     = tdb.get_hourly_df(contor, an=an)
            b1     = int(df["Total_B1"].sum()) if "Total_B1" in df.columns else 0
            b2     = int(df["Total_B2"].sum()) if "Total_B2" in df.columns else 0
            n_benzi = int(df["N_Benzi"].iloc[0]) if "N_Benzi" in df.columns else 2
            rezultate.append({
                "path":    path,
                "id":      contor,
                "randuri": len(df),
                "b1":      b1,
                "b2":      b2,
                "n_lanes": n_benzi,
            })
            try:
                update_centralizator(path, contor, CENTRAL_FILE_FOLDER)
            except Exception as e:
                print(f"[CENTRALIZATOR] Eroare [{contor}]: {e}")

    return rezultate
