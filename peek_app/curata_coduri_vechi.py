# =============================================================================
# curata_coduri_vechi.py — Curăță reziduurile codurilor vechi după rename
# =============================================================================
# Rulează DUPĂ rename_posturi.py.
# Pentru fiecare pereche din contor_alias:
#   - mută rândurile rămase din trafic_mzl (cod_vechi → cod_nou)
#   - șterge cod_vechi din contoare
# =============================================================================

import sqlite3

TRAFFIC_DB  = r"L:\BIDMRCT\datePEEK\SQLite\trafic.db"
CONTOARE_DB = r"L:\BIDMRCT\datePEEK\SQLite\contoare.db"


def curata():
    # Citim toate alias-urile înregistrate
    conn_c = sqlite3.connect(CONTOARE_DB)
    conn_c.row_factory = sqlite3.Row
    try:
        alias_rows = conn_c.execute(
            "SELECT cod_vechi, cod_nou FROM contor_alias"
        ).fetchall()
    except Exception:
        print("[INFO] Tabelul contor_alias nu există — nimic de făcut.")
        conn_c.close()
        return
    conn_c.close()

    if not alias_rows:
        print("[INFO] contor_alias e gol — nimic de curățat.")
        return

    print(f"Posturi de curățat: {len(alias_rows)}\n")

    for row in alias_rows:
        cod_vechi = row["cod_vechi"]
        cod_nou   = row["cod_nou"]
        print(f"{'─'*50}")
        print(f"  {cod_vechi} → {cod_nou}")

        # ── trafic.db: rânduri MZL rămase ────────────────────────────────────
        conn_t = sqlite3.connect(TRAFFIC_DB)
        try:
            conn_t.execute("PRAGMA journal_mode=WAL")

            # Șterge din cod_vechi lunile care există deja în cod_nou
            conn_t.execute("""
                DELETE FROM trafic_mzl
                WHERE contor = ?
                  AND (an, luna) IN (
                      SELECT an, luna FROM trafic_mzl WHERE contor = ?
                  )
            """, (cod_vechi, cod_nou))
            sterse = conn_t.execute("SELECT changes()").fetchone()[0]
            if sterse:
                print(f"  [trafic_mzl] {sterse} duplicat(e) șters(e) din {cod_vechi}.")

            # Mută ce a mai rămas
            conn_t.execute(
                "UPDATE trafic_mzl SET contor=? WHERE contor=?",
                (cod_nou, cod_vechi)
            )
            mutate = conn_t.execute("SELECT changes()").fetchone()[0]
            if mutate:
                print(f"  [trafic_mzl] {mutate} rând(uri) mutate {cod_vechi} → {cod_nou}.")
            else:
                print(f"  [trafic_mzl] Nimic rămas în {cod_vechi}.")

            conn_t.commit()
        except Exception as e:
            conn_t.rollback()
            print(f"  [EROARE trafic.db] {e}")
        finally:
            conn_t.close()

        # ── contoare.db: șterge codul vechi ──────────────────────────────────
        conn_c = sqlite3.connect(CONTOARE_DB)
        try:
            conn_c.execute("PRAGMA journal_mode=WAL")
            conn_c.execute("DELETE FROM contoare WHERE contor=?", (cod_vechi,))
            sters = conn_c.execute("SELECT changes()").fetchone()[0]
            if sters:
                print(f"  [contoare]   {cod_vechi} șters.")
            else:
                print(f"  [contoare]   {cod_vechi} nu era în tabel.")
            conn_c.commit()
        except Exception as e:
            conn_c.rollback()
            print(f"  [EROARE contoare.db] {e}")
        finally:
            conn_c.close()

    print(f"\n{'═'*50}")
    print("  Curățare finalizată.")
    print(f"{'═'*50}\n")


if __name__ == "__main__":
    curata()
