# =============================================================================
# GHID PRACTIC SQLite — PEEK/VEK Traffic Analyzer
# =============================================================================
# Autor: generat pentru Razvan / CESTRIN
# Scop:  ghid de utilizare zilnică a celor două baze de date SQLite
# =============================================================================


# ══════════════════════════════════════════════════════════════════════════════
# 1. STRUCTURA BAZELOR DE DATE
# ══════════════════════════════════════════════════════════════════════════════
#
#   L:\BIDMRCT\datePEEK\
#       trafic.db          ← date orare de trafic (toate contoarele, toți anii)
#       contoare.db        ← identificare contoare (drum, km, localitate, tip, IP)
#       0_Centralizator_PEEK-VEK.xlsx   ← generat din trafic.db la cerere
#
#   trafic.db conține 3 tabele:
#     • inregistrari_orare  — câte un rând per contor per oră
#     • mzl_manual          — suprascrieri MZL de la operatori
#     • fisiere_procesate   — log al fișierelor .bin/.log importate
#
#   contoare.db conține 1 tabel:
#     • contoare            — date de identificare (înlocuiește sheet-ul Contoare)
#
#   Legătura între ele: coloana "contor" (ex: "2168") prezentă în ambele DB.


# ══════════════════════════════════════════════════════════════════════════════
# 2. PRIMUL START — MIGRARE DATE EXISTENTE
# ══════════════════════════════════════════════════════════════════════════════
#
# Rulează o singură dată după instalarea versiunii cu SQLite:

from database import migrate_contoare_from_json, migrate_traffic_from_excel

# Migrare date identificare contoare din sheet-ul Excel existent
migrate_contoare_from_json()

# Migrare date orare din Excel-urile Raport_Clase_Peek deja generate
migrate_traffic_from_excel(r"L:\BIDMRCT\datePEEK", tip_sursa="PEEK")
migrate_traffic_from_excel(r"L:\BIDMRCT\datePEEK", tip_sursa="VEK")

# Verificare rezultat
from database import get_traffic_db
stats = get_traffic_db().stats()
print(f"Rânduri orare: {stats['randuri_orare']:,}")
print(f"Contoare:      {stats['contoare']}")
print(f"Dimensiune DB: {stats['size_mb']} MB")


# ══════════════════════════════════════════════════════════════════════════════
# 3. INTEROGĂRI FRECVENTE — COPIERE DIRECTĂ ÎN PYTHON
# ══════════════════════════════════════════════════════════════════════════════

import sqlite3
import pandas as pd

# Deschidere conexiune (read-only pentru interogări simple)
conn_t = sqlite3.connect(r"L:\BIDMRCT\datePEEK\trafic.db")
conn_c = sqlite3.connect(r"L:\BIDMRCT\datePEEK\contoare.db")


# ── 3.1 Ce contoare avem date în DB? ─────────────────────────────────────────

df = pd.read_sql("""
    SELECT contor,
           MIN(timestamp) AS prima_inregistrare,
           MAX(timestamp) AS ultima_inregistrare,
           COUNT(*)       AS ore_totale
    FROM inregistrari_orare
    GROUP BY contor
    ORDER BY contor
""", conn_t)
print(df)


# ── 3.2 Toate datele orare pentru contorul 2168 din 2024 ─────────────────────

df = pd.read_sql("""
    SELECT timestamp, total_b1, total_b2, total_general
    FROM inregistrari_orare
    WHERE contor = '2168' AND an = 2024
    ORDER BY timestamp
""", conn_t)


# ── 3.3 Media zilnică lunară calculată din DB (fără Excel) ───────────────────

df = pd.read_sql("""
    SELECT contor, an, luna,
           ROUND(AVG(total_general), 0) AS mzl_medie,
           COUNT(DISTINCT zi)           AS zile_cu_date
    FROM inregistrari_orare
    WHERE contor = '2168'
    GROUP BY contor, an, luna
    ORDER BY an, luna
""", conn_t)


# ── 3.4 Top 10 ore cu cel mai mult trafic (contor 2168, 2024) ────────────────

df = pd.read_sql("""
    SELECT timestamp, total_general
    FROM inregistrari_orare
    WHERE contor = '2168' AND an = 2024
    ORDER BY total_general DESC
    LIMIT 10
""", conn_t)


# ── 3.5 Zile cu Clasa_15 > 10% (posibile totalizatoare) ─────────────────────

df = pd.read_sql("""
    SELECT contor,
           DATE(timestamp) AS data,
           SUM(b1_clasa_15 + b2_clasa_15)        AS cls15_total,
           SUM(total_general)                     AS total,
           ROUND(
               100.0 * SUM(b1_clasa_15 + b2_clasa_15) / SUM(total_general),
               1
           ) AS pct_cls15
    FROM inregistrari_orare
    WHERE contor = '2168'
      AND total_general > 0
    GROUP BY contor, DATE(timestamp)
    HAVING pct_cls15 > 10
    ORDER BY data
""", conn_t)


# ── 3.6 Toate contoarele cu locația lor (JOIN între cele 2 DB) ───────────────
# SQLite nu face JOIN direct între fișiere diferite,
# dar putem folosi ATTACH:

conn_t.execute(f"ATTACH DATABASE '{r'L:\BIDMRCT\datePEEK\contoare.db'}' AS cdb")

df = pd.read_sql("""
    SELECT t.contor,
           c.drum,
           c.pozitie_km,
           c.localitate,
           MIN(t.timestamp) AS prima_data,
           MAX(t.timestamp) AS ultima_data,
           COUNT(*)         AS ore_totale
    FROM inregistrari_orare t
    LEFT JOIN cdb.contoare c ON t.contor = c.contor
    GROUP BY t.contor
    ORDER BY c.drum, t.contor
""", conn_t)
print(df.to_string())


# ── 3.7 Contoare fără date în luna curentă (posibil defecte) ─────────────────

import datetime
luna_curenta = datetime.date.today().month
an_curent    = datetime.date.today().year

df = pd.read_sql("""
    SELECT c.contor, c.drum, c.localitate
    FROM contoare c
    WHERE c.activ = 1
      AND c.contor NOT IN (
          SELECT DISTINCT contor
          FROM inregistrari_orare
          WHERE an = ? AND luna = ?
      )
    ORDER BY c.drum
""", conn_c, params=(an_curent, luna_curenta))
print("Contoare fără date luna curentă:")
print(df)


# ── 3.8 Suprascrieri MZL manual — audit complet ──────────────────────────────

df = pd.read_sql("""
    SELECT contor, an, luna, mzl_valoare,
           observatii, utilizator, modificat_la
    FROM mzl_manual
    ORDER BY modificat_la DESC
""", conn_t)


# ── 3.9 Fișiere procesate per contor ─────────────────────────────────────────

df = pd.read_sql("""
    SELECT contor, tip_sursa,
           COUNT(*)        AS nr_fisiere,
           SUM(inregistrari) AS total_ore,
           MAX(procesat_la)  AS ultima_procesare
    FROM fisiere_procesate
    GROUP BY contor, tip_sursa
    ORDER BY contor
""", conn_t)


# ══════════════════════════════════════════════════════════════════════════════
# 4. ACTUALIZARE DATE IDENTIFICARE CONTOARE
# ══════════════════════════════════════════════════════════════════════════════

from database import get_contoare_db

cdb = get_contoare_db()

# Adaugă/actualizează un contor
cdb.upsert("2168", {
    "Drum":       "DN1",
    "Pozitie_km": "km 45+200",
    "Localitate": "Ploiești",
    "Tip":        "ADR 3000 - Clasificator",
    "IP":         "192.168.1.100",
})

# Citește un contor
info = cdb.get("2168")
print(info)

# Toate contoarele active
toate = cdb.get_all()   # dict {"contor": {date...}}

# Dezactivează un contor (soft delete — datele de trafic rămân)
cdb.delete("9999")


# ══════════════════════════════════════════════════════════════════════════════
# 5. ADĂUGARE MZL MANUAL DIN COD
# ══════════════════════════════════════════════════════════════════════════════

from database import get_traffic_db

tdb = get_traffic_db()

# Suprascrie MZL pentru contorul 2168, luna august 2024
tdb.upsert_mzl_manual(
    contor="2168", an=2024, luna=8,
    valoare=1250.0,
    observatii="Contor reparat după defecțiune 10-15 aug",
    utilizator="Ion Popescu"
)

# Citește toate suprascrierile pentru un contor
mzl = tdb.get_mzl_manual("2168")
# → {(2024, 8): 1250.0, ...}


# ══════════════════════════════════════════════════════════════════════════════
# 6. GENERARE RAPOARTE DIN DB (fără a reporni parsarea)
# ══════════════════════════════════════════════════════════════════════════════

from db_report import generate_report_from_db, generate_centralizator_from_db

# Raport pentru un contor, un an
path = generate_report_from_db("2168", an=2024,
                                output_dir=r"L:\BIDMRCT\datePEEK")
print(f"Excel generat: {path}")

# Regenerează centralizatorul complet din DB
generate_centralizator_from_db()


# ══════════════════════════════════════════════════════════════════════════════
# 7. BACKUP ȘI ÎNTREȚINERE
# ══════════════════════════════════════════════════════════════════════════════

import shutil, datetime

# Backup simplu (SQLite e un singur fișier — copierea e suficientă)
data = datetime.date.today().strftime("%Y%m%d")
shutil.copy(r"L:\BIDMRCT\datePEEK\trafic.db",
            rf"L:\BIDMRCT\backup\trafic_{data}.db")
shutil.copy(r"L:\BIDMRCT\datePEEK\contoare.db",
            rf"L:\BIDMRCT\backup\contoare_{data}.db")

# Optimizare periodică (lunar) — reconstruiește indecșii, eliberează spațiu
conn = sqlite3.connect(r"L:\BIDMRCT\datePEEK\trafic.db")
conn.execute("VACUUM")      # reconstruiește fișierul (durează câteva minute)
conn.execute("ANALYZE")     # actualizează statisticile pentru query planner
conn.close()

# Verificare integritate
conn = sqlite3.connect(r"L:\BIDMRCT\datePEEK\trafic.db")
result = conn.execute("PRAGMA integrity_check").fetchone()[0]
print(f"Integritate DB: {result}")   # ar trebui să fie "ok"
conn.close()


# ══════════════════════════════════════════════════════════════════════════════
# 8. INSTRUMENTE GRAFICE RECOMANDATE (opțional)
# ══════════════════════════════════════════════════════════════════════════════
#
# DB Browser for SQLite (gratuit, Windows)
#   https://sqlitebrowser.org/
#   → interfață vizuală pentru browse date, rulat interogări SQL, export CSV
#   → util pentru operatori care nu scriu cod Python
#
# DBeaver Community (gratuit)
#   https://dbeaver.io/
#   → mai avansat, suportă și migrare viitoare la PostgreSQL fără schimbări SQL
#
# În Python: Jupyter Notebook + pandas
#   pip install notebook
#   jupyter notebook
#   → interogare interactivă, grafice matplotlib/plotly direct din DB
