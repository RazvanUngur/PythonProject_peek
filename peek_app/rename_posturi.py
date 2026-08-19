# =============================================================================
# rename_posturi.py — Migrare coduri post după reorganizare DRDP
# =============================================================================
# Mută istoricul de trafic din codul vechi în codul nou pentru mai multe
# posturi simultan. Perioadele de date nu se suprapun între coduri.
#
# UTILIZARE:
#   1. Editează lista REDENUMIRI de mai jos cu perechile tale reale
#   2. Verifică TRAFFIC_DB și CONTOARE_DB să pointeze la fișierele corecte
#   3. Fă BACKUP la ambele .db înainte de rulare
#   4. Rulează: python rename_posturi.py
# =============================================================================

import sqlite3
from datetime import date

# ── Căi baze de date ──────────────────────────────────────────────────────────
TRAFFIC_DB  = r"L:\BIDMRCT\datePEEK\SQLite\trafic.db"
CONTOARE_DB = r"L:\BIDMRCT\datePEEK\SQLite\contoare.db"

# ── Lista redenumiri: (cod_vechi, cod_nou) ────────────────────────────────────
# Editează această listă cu perechile tale reale:
REDENUMIRI = [
    ("8203", "1203"),
    ("7642", "8642"),
    ("7598", "8598"),
    ("1044", "8044"),
    ("3848", "3343"),

    # adaugă câte rânduri ai nevoie...
]

# ── Motiv comun pentru toate redenumirile ─────────────────────────────────────
MOTIV = "Reorganizare DRDP 2023"


# =============================================================================
# Funcții interne
# =============================================================================

def _verifica_suprapunere(conn, cod_vechi: str, cod_nou: str) -> list:
    """
    Returnează lista de (an, luna, zi, ora) care se suprapun între cele două coduri.
    Verificare la nivel de oră — în luna de tranziție e normal ca ambele
    coduri să aibă date, dar pe ore diferite.
    """
    rows = conn.execute("""
        SELECT an, luna, zi, ora FROM inregistrari_orare WHERE contor = ?
        INTERSECT
        SELECT an, luna, zi, ora FROM inregistrari_orare WHERE contor = ?
    """, (cod_vechi, cod_nou)).fetchall()
    return [(r[0], r[1], r[2], r[3]) for r in rows]


def _ensure_alias_table(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS contor_alias (
            cod_vechi   TEXT PRIMARY KEY,
            cod_nou     TEXT NOT NULL,
            data_schimb TEXT,
            motiv       TEXT
        )
    """)


def rename_post(cod_vechi: str, cod_nou: str, motiv: str = "") -> bool:
    """
    Mută toate datele din cod_vechi în cod_nou în trafic.db și contoare.db.
    Returnează True dacă a reușit, False dacă a fost oprită din cauza unei erori.
    """
    print(f"\n{'─'*60}")
    print(f"  Procesare: {cod_vechi} → {cod_nou}")
    print(f"{'─'*60}")

    # ── trafic.db ─────────────────────────────────────────────────────────────
    conn_t = sqlite3.connect(TRAFFIC_DB)
    conn_t.row_factory = sqlite3.Row
    try:
        # Verificare suprapunere
        overlap = _verifica_suprapunere(conn_t, cod_vechi, cod_nou)
        if overlap:
            print(f"  [EROARE] Suprapunere date la nivel de oră ({len(overlap)} conflicte):")
            for an, luna, zi, ora in overlap[:10]:  # afișăm max 10 exemple
                print(f"           {zi:02d}.{luna:02d}.{an} ora {ora}")
            if len(overlap) > 10:
                print(f"           ... și încă {len(overlap)-10} conflicte.")
            print(f"  [SKIP]   {cod_vechi} → {cod_nou} anulat.")
            return False

        # Verificare că există date de mutat
        n_orare = conn_t.execute(
            "SELECT COUNT(*) FROM inregistrari_orare WHERE contor=?",
            (cod_vechi,)
        ).fetchone()[0]
        n_mzl = conn_t.execute(
            "SELECT COUNT(*) FROM trafic_mzl WHERE contor=?",
            (cod_vechi,)
        ).fetchone()[0]

        if n_orare == 0 and n_mzl == 0:
            print(f"  [INFO]   {cod_vechi} nu are date în trafic.db — skip trafic.")
        else:
            conn_t.execute("PRAGMA journal_mode=WAL")
            conn_t.execute(
                "UPDATE inregistrari_orare SET contor=? WHERE contor=?",
                (cod_nou, cod_vechi)
            )
            # trafic_mzl — șterge rândurile din cod_vechi care au același
            # (an, luna) ca în cod_nou; păstrăm valorile din cod_nou (mai recente)
            conn_t.execute("""
                DELETE FROM trafic_mzl
                WHERE contor = ?
                  AND (an, luna) IN (
                      SELECT an, luna FROM trafic_mzl WHERE contor = ?
                  )
            """, (cod_vechi, cod_nou))
            n_mzl_sterse = conn_t.execute("SELECT changes()").fetchone()[0]
            if n_mzl_sterse:
                print(f"  [INFO]   trafic_mzl: {n_mzl_sterse} duplicat(e) din "
                      f"{cod_vechi} șterse (se păstrează valorile din {cod_nou}).")

            conn_t.execute(
                "UPDATE trafic_mzl SET contor=? WHERE contor=?",
                (cod_nou, cod_vechi)
            )
            conn_t.commit()
            print(f"  [OK]     trafic.db: {n_orare} înreg. orare + {n_mzl} MZL mutate.")

    except Exception as e:
        conn_t.rollback()
        print(f"  [EROARE] trafic.db: {e}")
        return False
    finally:
        conn_t.close()

    # ── contoare.db ───────────────────────────────────────────────────────────
    conn_c = sqlite3.connect(CONTOARE_DB)
    conn_c.row_factory = sqlite3.Row
    try:
        conn_c.execute("PRAGMA journal_mode=WAL")

        row_vechi = conn_c.execute(
            "SELECT * FROM contoare WHERE contor=?", (cod_vechi,)
        ).fetchone()
        row_nou = conn_c.execute(
            "SELECT * FROM contoare WHERE contor=?", (cod_nou,)
        ).fetchone()

        if row_vechi and row_nou:
            # Ambele coduri există în contoare.db:
            # completează câmpurile goale din cod_nou cu valorile din cod_vechi
            # (coordonatele și metadatele de la postul fizic rămân)
            conn_c.execute("""
                UPDATE contoare SET
                    lat        = COALESCE(NULLIF(lat,        ''), (SELECT lat        FROM contoare WHERE contor=?)),
                    lng        = COALESCE(NULLIF(lng,        ''), (SELECT lng        FROM contoare WHERE contor=?)),
                    drum       = COALESCE(NULLIF(drum,       ''), (SELECT drum       FROM contoare WHERE contor=?)),
                    pozitie_km = COALESCE(NULLIF(pozitie_km, ''), (SELECT pozitie_km FROM contoare WHERE contor=?)),
                    localitate = COALESCE(NULLIF(localitate, ''), (SELECT localitate FROM contoare WHERE contor=?)),
                    tip        = COALESCE(NULLIF(tip,        ''), (SELECT tip        FROM contoare WHERE contor=?)),
                    ip         = COALESCE(NULLIF(ip,         ''), (SELECT ip         FROM contoare WHERE contor=?))
                WHERE contor=?
            """, (cod_vechi,) * 7 + (cod_nou,))
            conn_c.execute("DELETE FROM contoare WHERE contor=?", (cod_vechi,))
            print(f"  [OK]     contoare.db: metadate îmbinate, {cod_vechi} șters.")

        elif row_vechi and not row_nou:
            # Doar codul vechi există — redenumire directă
            conn_c.execute(
                "UPDATE contoare SET contor=? WHERE contor=?",
                (cod_nou, cod_vechi)
            )
            print(f"  [OK]     contoare.db: {cod_vechi} redenumit în {cod_nou}.")

        elif not row_vechi and row_nou:
            print(f"  [INFO]   contoare.db: {cod_nou} există deja, {cod_vechi} absent — nimic de făcut.")

        else:
            print(f"  [INFO]   contoare.db: niciun cod găsit — nimic de făcut.")

        # Înregistrare audit
        _ensure_alias_table(conn_c)
        conn_c.execute("""
            INSERT OR REPLACE INTO contor_alias (cod_vechi, cod_nou, data_schimb, motiv)
            VALUES (?, ?, ?, ?)
        """, (cod_vechi, cod_nou, date.today().isoformat(), motiv))
        conn_c.commit()
        print(f"  [OK]     Alias înregistrat în contor_alias.")

    except Exception as e:
        conn_c.rollback()
        print(f"  [EROARE] contoare.db: {e}")
        return False
    finally:
        conn_c.close()

    return True


# =============================================================================
# Entry point
# =============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("  MIGRARE CODURI POST — Reorganizare DRDP")
    print(f"  Data: {date.today().isoformat()}")
    print(f"  Motiv: {MOTIV}")
    print("=" * 60)
    print(f"\n  Posturi de procesat: {len(REDENUMIRI)}")

    ok_list   = []
    skip_list = []

    for cod_vechi, cod_nou in REDENUMIRI:
        success = rename_post(cod_vechi, cod_nou, motiv=MOTIV)
        if success:
            ok_list.append((cod_vechi, cod_nou))
        else:
            skip_list.append((cod_vechi, cod_nou))

    # ── Sumar final ───────────────────────────────────────────────────────────
    print(f"\n{'═'*60}")
    print(f"  SUMAR FINAL")
    print(f"{'═'*60}")
    print(f"  Reușite : {len(ok_list)}")
    for v, n in ok_list:
        print(f"    ✓  {v} → {n}")
    if skip_list:
        print(f"  Eșuate / Skip : {len(skip_list)}")
        for v, n in skip_list:
            print(f"    ✗  {v} → {n}")
    print()
