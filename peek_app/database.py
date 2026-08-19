# =============================================================================
# database.py — Gestionare baze de date SQLite
# =============================================================================
#
# Două baze de date separate care lucrează împreună:
#
#   trafic.db   — date orare de trafic de la toate contoarele
#   contoare.db — date de identificare contoare (drum, km, localitate, tip, IP, X, Y)
#
# Exportă:
#   TrafficDB, ContoareDB
#   get_traffic_db()   → instanță singleton TrafficDB
#   get_contoare_db()  → instanță singleton ContoareDB
# =============================================================================

import os
import sqlite3
import threading
import pandas as pd
from datetime import datetime

from config import CENTRAL_FILE_FOLDER, SQLITE_FOLDER

# ── Căi fișiere DB ────────────────────────────────────────────────────────────
DB_FOLDER   = os.path.join(CENTRAL_FILE_FOLDER, SQLITE_FOLDER)
TRAFFIC_DB  = os.path.join(DB_FOLDER, "trafic.db")
CONTOARE_DB = os.path.join(DB_FOLDER, "contoare.db")

# ── Singletons ────────────────────────────────────────────────────────────────
_traffic_instance  = None
_contoare_instance = None
_lock = threading.Lock()

# =============================================================================
# SCHEMA — trafic.db
# =============================================================================

_TRAFFIC_SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS inregistrari_orare (
    contor          TEXT    NOT NULL,
    timestamp       TEXT    NOT NULL,
    an              INTEGER NOT NULL,
    luna            INTEGER NOT NULL,
    zi              INTEGER NOT NULL,
    ora             INTEGER NOT NULL,

    b1_clasa_1      INTEGER DEFAULT 0,
    b1_clasa_2      INTEGER DEFAULT 0,
    b1_clasa_3      INTEGER DEFAULT 0,
    b1_clasa_4      INTEGER DEFAULT 0,
    b1_clasa_5      INTEGER DEFAULT 0,
    b1_clasa_6      INTEGER DEFAULT 0,
    b1_clasa_7      INTEGER DEFAULT 0,
    b1_clasa_8      INTEGER DEFAULT 0,
    b1_clasa_15     INTEGER DEFAULT 0,
    total_b1        INTEGER DEFAULT 0,

    b2_clasa_1      INTEGER DEFAULT 0,
    b2_clasa_2      INTEGER DEFAULT 0,
    b2_clasa_3      INTEGER DEFAULT 0,
    b2_clasa_4      INTEGER DEFAULT 0,
    b2_clasa_5      INTEGER DEFAULT 0,
    b2_clasa_6      INTEGER DEFAULT 0,
    b2_clasa_7      INTEGER DEFAULT 0,
    b2_clasa_8      INTEGER DEFAULT 0,
    b2_clasa_15     INTEGER DEFAULT 0,
    total_b2        INTEGER DEFAULT 0,

    b3_clasa_1      INTEGER DEFAULT 0,
    b3_clasa_2      INTEGER DEFAULT 0,
    b3_clasa_3      INTEGER DEFAULT 0,
    b3_clasa_4      INTEGER DEFAULT 0,
    b3_clasa_5      INTEGER DEFAULT 0,
    b3_clasa_6      INTEGER DEFAULT 0,
    b3_clasa_7      INTEGER DEFAULT 0,
    b3_clasa_8      INTEGER DEFAULT 0,
    b3_clasa_15     INTEGER DEFAULT 0,
    total_b3        INTEGER DEFAULT 0,

    b4_clasa_1      INTEGER DEFAULT 0,
    b4_clasa_2      INTEGER DEFAULT 0,
    b4_clasa_3      INTEGER DEFAULT 0,
    b4_clasa_4      INTEGER DEFAULT 0,
    b4_clasa_5      INTEGER DEFAULT 0,
    b4_clasa_6      INTEGER DEFAULT 0,
    b4_clasa_7      INTEGER DEFAULT 0,
    b4_clasa_8      INTEGER DEFAULT 0,
    b4_clasa_15     INTEGER DEFAULT 0,
    total_b4        INTEGER DEFAULT 0,

    b5_clasa_1      INTEGER DEFAULT 0,
    b5_clasa_2      INTEGER DEFAULT 0,
    b5_clasa_3      INTEGER DEFAULT 0,
    b5_clasa_4      INTEGER DEFAULT 0,
    b5_clasa_5      INTEGER DEFAULT 0,
    b5_clasa_6      INTEGER DEFAULT 0,
    b5_clasa_7      INTEGER DEFAULT 0,
    b5_clasa_8      INTEGER DEFAULT 0,
    b5_clasa_15     INTEGER DEFAULT 0,
    total_b5        INTEGER DEFAULT 0,

    b6_clasa_1      INTEGER DEFAULT 0,
    b6_clasa_2      INTEGER DEFAULT 0,
    b6_clasa_3      INTEGER DEFAULT 0,
    b6_clasa_4      INTEGER DEFAULT 0,
    b6_clasa_5      INTEGER DEFAULT 0,
    b6_clasa_6      INTEGER DEFAULT 0,
    b6_clasa_7      INTEGER DEFAULT 0,
    b6_clasa_8      INTEGER DEFAULT 0,
    b6_clasa_15     INTEGER DEFAULT 0,
    total_b6        INTEGER DEFAULT 0,

    total_general   INTEGER DEFAULT 0,
    n_benzi         INTEGER DEFAULT 2,
    tip_sursa       TEXT    DEFAULT 'PEEK',
    source_file     TEXT    DEFAULT '',

    PRIMARY KEY (contor, timestamp)
);

CREATE INDEX IF NOT EXISTS idx_trafic_contor
    ON inregistrari_orare (contor);
CREATE INDEX IF NOT EXISTS idx_trafic_contor_an_luna
    ON inregistrari_orare (contor, an, luna);
CREATE INDEX IF NOT EXISTS idx_trafic_an_luna
    ON inregistrari_orare (an, luna);

CREATE TABLE IF NOT EXISTS mzl_manual (
    contor          TEXT    NOT NULL,
    an              INTEGER NOT NULL,
    luna            INTEGER NOT NULL,
    mzl_valoare     REAL    NOT NULL,
    observatii      TEXT    DEFAULT '',
    utilizator      TEXT    DEFAULT '',
    modificat_la    TEXT    DEFAULT '',

    PRIMARY KEY (contor, an, luna)
);

CREATE INDEX IF NOT EXISTS idx_mzl_contor
    ON mzl_manual (contor);

CREATE TABLE IF NOT EXISTS trafic_mzl (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    contor          TEXT    NOT NULL,
    an              INTEGER NOT NULL,
    luna            INTEGER NOT NULL,
    mzl_calculat    REAL    NOT NULL,
    mzl_final       REAL    NOT NULL,
    este_manual     INTEGER DEFAULT 0,
    indicator       TEXT    DEFAULT '',
    zile_valide     INTEGER DEFAULT 0,
    zile_luna       INTEGER DEFAULT 0,
    calculat_la     TEXT    DEFAULT '',

    UNIQUE (contor, an, luna)
);

CREATE INDEX IF NOT EXISTS idx_trafic_mzl_contor
    ON trafic_mzl (contor);
CREATE INDEX IF NOT EXISTS idx_trafic_mzl_contor_an
    ON trafic_mzl (contor, an);

CREATE TABLE IF NOT EXISTS fisiere_procesate (
    cale_fisier     TEXT    PRIMARY KEY,
    contor          TEXT    NOT NULL,
    tip_sursa       TEXT    NOT NULL,
    inregistrari    INTEGER DEFAULT 0,
    procesat_la     TEXT    NOT NULL,
    checksum        TEXT    DEFAULT '',
    perioada_min    TEXT    DEFAULT '',
    perioada_max    TEXT    DEFAULT ''
);
"""

# =============================================================================
# SCHEMA — contoare.db
# =============================================================================

_CONTOARE_SCHEMA = """
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS contoare (
    contor          TEXT    PRIMARY KEY,
    drum            TEXT    DEFAULT '',
    pozitie_km      TEXT    DEFAULT '',
    localitate      TEXT    DEFAULT '',
    tip             TEXT    DEFAULT '',
    ip              TEXT    DEFAULT '',
    x               REAL    DEFAULT NULL,
    y               REAL    DEFAULT NULL,
    lat             REAL    DEFAULT NULL,
    lng             REAL    DEFAULT NULL,
    drdp            TEXT    DEFAULT '',
    activ           INTEGER DEFAULT 1,
    adaugat_la      TEXT    DEFAULT '',
    modificat_la    TEXT    DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_contoare_drum
    ON contoare (drum);
CREATE INDEX IF NOT EXISTS idx_contoare_localitate
    ON contoare (localitate);
"""


# =============================================================================
# CLASS TrafficDB
# =============================================================================

class TrafficDB:
    def __init__(self, db_path: str = TRAFFIC_DB):
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self.db_path = db_path
        self._local  = threading.local()
        self._init_schema()

    def _conn(self) -> sqlite3.Connection:
        conn = getattr(self._local, "conn", None)
        # Reconectare dacă conexiunea e invalidă sau a fost închisă de alt proces
        if conn is not None:
            try:
                conn.execute("SELECT 1")
            except Exception:
                conn = None
        if conn is None:
            conn = sqlite3.connect(
                self.db_path,
                timeout=15,               # Asteapta max 15s la lock extern
                check_same_thread=False,
            )
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode = WAL")
            conn.execute("PRAGMA synchronous = NORMAL")
            conn.execute("PRAGMA cache_size = -32000")
            conn.execute("PRAGMA busy_timeout = 10000")  # Retry automat 10s la lock
            conn.execute("PRAGMA wal_autocheckpoint = 100")
            self._local.conn = conn
        return conn

    def _init_schema(self):
        conn = self._conn()
        conn.executescript(_TRAFFIC_SCHEMA)
        conn.commit()
        self._migrate_schema(conn)

    def _migrate_schema(self, conn):
        """
        Migrări aditive pentru DB-uri create înainte de introducerea unor
        coloane noi. SQLite nu suportă `ADD COLUMN IF NOT EXISTS`, deci
        verificăm manual prin PRAGMA table_info.
        """
        try:
            cols = {r["name"] for r in conn.execute(
                "PRAGMA table_info(fisiere_procesate)").fetchall()}
            if "perioada_min" not in cols:
                conn.execute(
                    "ALTER TABLE fisiere_procesate ADD COLUMN perioada_min TEXT DEFAULT ''")
            if "perioada_max" not in cols:
                conn.execute(
                    "ALTER TABLE fisiere_procesate ADD COLUMN perioada_max TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            pass  # migrare best-effort; nu blocăm pornirea aplicației

    @staticmethod
    def _resolve_contor_alias(site_id: str) -> str:
        """
        Verifică dacă postul site_id a fost redenumit (există în contor_alias).
        Dacă da, returnează codul nou; altfel returnează site_id nemodificat.
        Astfel, dacă se reprocesează fișiere vechi cu codul 8203, datele
        vor fi salvate automat sub codul nou 1203.
        """
        try:
            conn = sqlite3.connect(CONTOARE_DB, timeout=10)
            try:
                row = conn.execute(
                    "SELECT cod_nou FROM contor_alias WHERE cod_vechi = ?",
                    (site_id,)
                ).fetchone()
                if row:
                    print(f"  [ALIAS] {site_id} → {row[0]} (redenumit)")
                    return row[0]
            finally:
                conn.close()
        except Exception:
            pass  # dacă tabelul nu există încă, nu facem nimic
        return site_id

    def _execute_with_retry(self, sql: str, params=(), retries: int = 5,
                             delay: float = 0.5) -> sqlite3.Cursor:
        """
        Execută o interogare SQL cu retry automat la OperationalError (database locked).
        Util pentru scrieri concurente multi-utilizator pe rețea.
        """
        import time
        last_err = None
        for attempt in range(retries):
            try:
                conn = self._conn()
                cur = conn.execute(sql, params)
                conn.commit()
                return cur
            except sqlite3.OperationalError as e:
                last_err = e
                if "locked" in str(e).lower():
                    # Invalidam conexiunea ca sa fortam reconectare curata
                    try:
                        self._local.conn.close()
                    except Exception:
                        pass
                    self._local.conn = None
                    time.sleep(delay * (attempt + 1))
                else:
                    raise
        raise sqlite3.OperationalError(
            f"DB locked dupa {retries} incercari: {last_err}")


    def upsert_hourly_df(self, df: pd.DataFrame, site_id: str,
                         tip_sursa: str = "PEEK",
                         source_files: list = None,
                         source_file_periods: dict = None) -> int:
        """
        Inserează/actualizează date orare dintr-un DataFrame pandas.
        Returnează numărul de rânduri scrise.

        source_file_periods: dict opțional {basename → (perioada_min, perioada_max)}
        cu perioada REALĂ a fiecărui fișier individual (calculată de parser
        din conținutul lui brut, înainte de orice concatenare/deduplicare pe
        contor). Dacă e furnizat, se salvează per-fișier în `fisiere_procesate`,
        astfel încât fiecare fișier procesat vreodată să-și păstreze propria
        perioadă vizibilă, indiferent dacă a fost procesat singur sau
        împreună cu alte fișiere pentru același contor (caz în care coloana
        `source_file` din `inregistrari_orare` e comună pentru tot lotul).
        """
        if df is None or df.empty:
            return 0

        # Dacă postul a fost redenumit, folosim automat codul nou
        site_id = self._resolve_contor_alias(site_id)

        df = df.copy()

        # Normalizăm Timestamp
        if "Timestamp" in df.columns:
            df["_ts"] = pd.to_datetime(df["Timestamp"], errors="coerce")
        elif "Data_Ora" in df.columns:
            df["_ts"] = pd.to_datetime(df["Data_Ora"],
                                       format="%d.%m.%Y %H:%M", errors="coerce")
        else:
            raise ValueError("DataFrame nu are coloana Timestamp sau Data_Ora")

        df = df.dropna(subset=["_ts"])
        if df.empty:
            return 0

        source_file_str = (
            ", ".join(os.path.basename(f) for f in source_files)
            if source_files else ""
        )

        n_benzi = int(df["N_Benzi"].iloc[0]) if "N_Benzi" in df.columns else 2

        rows = []
        for _, row in df.iterrows():
            ts = row["_ts"]
            r = {
                "contor":    str(site_id),
                "timestamp": ts.strftime("%Y-%m-%d %H:%M:%S"),
                "an":        int(ts.year),
                "luna":      int(ts.month),
                "zi":        int(ts.day),
                "ora":       int(ts.hour),
                "n_benzi":   n_benzi,
                "tip_sursa": tip_sursa,
                "source_file": source_file_str,
                "total_general": int(row.get("Total_General", 0)),
            }
            for b in range(1, 7):
                for cls in list(range(1, 9)) + [15]:
                    col = f"B{b}_Clasa_{cls}"
                    r[f"b{b}_clasa_{cls}"] = int(row.get(col, 0))
                r[f"total_b{b}"] = int(row.get(f"Total_B{b}", 0))
            rows.append(r)

        if not rows:
            return 0

        conn = self._conn()
        cols = list(rows[0].keys())
        placeholders = ", ".join(f":{c}" for c in cols)
        sql = (f"INSERT OR REPLACE INTO inregistrari_orare "
               f"({', '.join(cols)}) VALUES ({placeholders})")

        conn.executemany(sql, rows)
        conn.commit()

        # Log fișiere procesate — fiecare fișier își păstrează PROPRIA
        # perioadă (din source_file_periods), nu perioada agregată a lotului.
        if source_files:
            for fp in source_files:
                fname = os.path.basename(fp)
                p_min = p_max = ""
                if source_file_periods and fname in source_file_periods:
                    p_min, p_max = source_file_periods[fname]
                conn.execute("""
                    INSERT OR REPLACE INTO fisiere_procesate
                    (cale_fisier, contor, tip_sursa, inregistrari, procesat_la,
                     perioada_min, perioada_max)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    os.path.abspath(fp), str(site_id), tip_sursa,
                    len(rows), datetime.now().isoformat(timespec="seconds"),
                    p_min, p_max
                ))
            conn.commit()

        return len(rows)

    # ── Citire date ───────────────────────────────────────────────────────────

    def get_hourly_df(self, contor: str,
                      an: int = None, luna: int = None) -> pd.DataFrame:
        """Returnează DataFrame orar pentru un contor."""
        params = [contor]
        where  = "contor = ?"
        if an is not None:
            where += " AND an = ?"; params.append(an)
        if luna is not None:
            where += " AND luna = ?"; params.append(luna)

        sql = f"SELECT * FROM inregistrari_orare WHERE {where} ORDER BY timestamp"
        df = pd.read_sql_query(sql, self._conn(), params=params)
        if df.empty:
            return df

        rename = {}
        for b in range(1, 7):
            for cls in list(range(1, 9)) + [15]:
                rename[f"b{b}_clasa_{cls}"] = f"B{b}_Clasa_{cls}"
            rename[f"total_b{b}"] = f"Total_B{b}"
        rename["total_general"] = "Total_General"
        rename["n_benzi"]       = "N_Benzi"
        df = df.rename(columns=rename)

        df["_ts"] = pd.to_datetime(df["timestamp"])
        df["Data_Ora"] = df["_ts"].dt.strftime("%d.%m.%Y %H:%M")
        df["Timestamp"] = df["_ts"]
        df["Contor"] = df["contor"]
        df = df.drop(columns=["_ts", "contor", "an", "luna", "zi", "ora",
                               "tip_sursa", "source_file"], errors="ignore")
        return df

    def get_contoare_disponibile(self) -> list:
        rows = self._conn().execute(
            "SELECT DISTINCT contor FROM inregistrari_orare ORDER BY contor"
        ).fetchall()
        return [r["contor"] for r in rows]

    def get_ani_disponibili(self, contor: str) -> list:
        rows = self._conn().execute(
            "SELECT DISTINCT an FROM inregistrari_orare "
            "WHERE contor = ? ORDER BY an", (contor,)
        ).fetchall()
        return [r["an"] for r in rows]

    def get_luni_disponibile(self, contor: str, an: int) -> list:
        rows = self._conn().execute(
            "SELECT DISTINCT luna FROM inregistrari_orare "
            "WHERE contor = ? AND an = ? ORDER BY luna", (contor, an)
        ).fetchall()
        return [r["luna"] for r in rows]

    def get_source_files(self, contor: str) -> list:
        """
        Returnează lista căilor complete ale fișierelor sursă procesate
        pentru un contor, în ordinea procesării.

        Notă: `fisiere_procesate` are cheie primară pe calea completă, deci
        dacă un fișier e mutat în alt folder și reprocesat, apare aici de
        două ori (calea veche + calea nouă) — comportament intenționat,
        pentru a păstra vizibil istoricul complet. Vezi `get_source_files_periods()`
        pentru varianta cu perioada de date aferentă fiecărui fișier, folosită
        de raportul "Fișiere sursă" din Excel.
        """
        rows = self._conn().execute(
            "SELECT cale_fisier FROM fisiere_procesate "
            "WHERE contor = ? ORDER BY procesat_la",
            (contor,)
        ).fetchall()
        return [r["cale_fisier"] for r in rows]

    def get_source_files_periods(self, contor: str) -> list:
        """
        Returnează fișierele sursă procesate pentru un contor, cu perioada
        proprie fiecăruia (salvată direct la procesare, în
        `fisiere_procesate.perioada_min/max` — vezi `upsert_hourly_df`).

        Deduplicare pe NUME de fișier (basename): dacă același fișier
        .bin/.log a fost procesat de mai multe ori din căi diferite (mutat,
        redenumit folderul, reprocesat manual etc.), se păstrează DOAR
        ultima cale procesată (după `procesat_la`), cu perioada ei —
        variantele vechi ale aceluiași fișier nu mai apar în listă.

        Returnează o listă de dict-uri, sortată cronologic după începutul
        perioadei (fișierele fără perioadă cunoscută apar la final):
            {
                "fisiere": [basename],
                "cai":     [cale_completă],   # ultima cale cunoscută
                "perioada_min": "dd.mm.yyyy" | "",
                "perioada_max": "dd.mm.yyyy" | "",
                "n_inregistrari": int,
            }
        """
        rows = self._conn().execute("""
            SELECT cale_fisier, inregistrari, perioada_min, perioada_max, procesat_la
            FROM fisiere_procesate
            WHERE contor = ?
        """, (contor,)).fetchall()

        if not rows:
            return []

        # Păstrăm doar ultima procesare (procesat_la cel mai recent) per
        # basename — procesat_la e ISO ("YYYY-MM-DDTHH:MM:SS"), deci
        # comparabil direct ca string.
        latest_by_basename = {}
        for r in rows:
            bn = os.path.basename(r["cale_fisier"])
            prev = latest_by_basename.get(bn)
            if prev is None or (r["procesat_la"] or "") > (prev["procesat_la"] or ""):
                latest_by_basename[bn] = r

        def _sort_key(p_min):
            d = self._parse_ro_date_safe(p_min)
            from datetime import date as _date
            return (d is None, d or _date.max)

        result = [{
            "fisiere":        [os.path.basename(r["cale_fisier"])],
            "cai":            [r["cale_fisier"]],
            "perioada_min":   r["perioada_min"] or "",
            "perioada_max":   r["perioada_max"] or "",
            "n_inregistrari": r["inregistrari"],
        } for r in latest_by_basename.values()]

        result.sort(key=lambda x: _sort_key(x["perioada_min"]))
        return result

    @staticmethod
    def _parse_ro_date_safe(s):
        try:
            return datetime.strptime(str(s).strip(), "%d.%m.%Y").date()
        except Exception:
            return None

        return result

    def fisier_procesat(self, filepath: str) -> bool:
        row = self._conn().execute(
            "SELECT cale_fisier FROM fisiere_procesate WHERE cale_fisier = ?",
            (os.path.abspath(filepath),)
        ).fetchone()
        return row is not None

    # ── MZL manual ────────────────────────────────────────────────────────────

    def upsert_mzl_manual(self, contor: str, an: int, luna: int,
                          valoare: float, observatii: str = "",
                          utilizator: str = "") -> None:
        """Salvează MZL manual cu retry automat la lock multi-utilizator."""
        self._execute_with_retry("""
            INSERT OR REPLACE INTO mzl_manual
            (contor, an, luna, mzl_valoare, observatii, utilizator, modificat_la)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (contor, an, luna, valoare, observatii, utilizator,
              datetime.now().isoformat(timespec="seconds")))

    def get_mzl_manual(self, contor: str) -> dict:
        """Returnează dict {(an, luna): valoare} pentru un contor.
        Reconectează automat dacă conexiunea e invalidă (ex: după lock extern)."""
        try:
            rows = self._conn().execute(
                "SELECT an, luna, mzl_valoare FROM mzl_manual WHERE contor = ?",
                (contor,)
            ).fetchall()
        except sqlite3.OperationalError:
            # Forțăm reconectare și reîncercăm
            self._local.conn = None
            rows = self._conn().execute(
                "SELECT an, luna, mzl_valoare FROM mzl_manual WHERE contor = ?",
                (contor,)
            ).fetchall()
        return {(r["an"], r["luna"]): r["mzl_valoare"] for r in rows}

    def get_all_mzl_manual(self) -> pd.DataFrame:
        return pd.read_sql_query(
            "SELECT * FROM mzl_manual ORDER BY contor, an, luna",
            self._conn()
        )

    # ── trafic_mzl ────────────────────────────────────────────────────────────

    def upsert_trafic_mzl(self, contor: str, an: int, luna: int,
                           mzl_calculat: float, mzl_final: float,
                           este_manual: int = 0, indicator: str = "",
                           zile_valide: int = 0, zile_luna: int = 0) -> None:
        """
        Inserează/actualizează MZL final în trafic_mzl.
        mzl_calculat = valoarea calculată automat din Date prelucrate (înainte de override manual).
        mzl_final    = valoarea folosită efectiv în raport (după override manual dacă există).
        este_manual  = 1 dacă mzl_final provine din mzl_manual, 0 dacă e calculat automat.
        """
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._execute_with_retry("""
            INSERT INTO trafic_mzl
                (contor, an, luna, mzl_calculat, mzl_final,
                 este_manual, indicator, zile_valide, zile_luna, calculat_la)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(contor, an, luna) DO UPDATE SET
                mzl_calculat = excluded.mzl_calculat,
                mzl_final    = excluded.mzl_final,
                este_manual  = excluded.este_manual,
                indicator    = excluded.indicator,
                zile_valide  = excluded.zile_valide,
                zile_luna    = excluded.zile_luna,
                calculat_la  = excluded.calculat_la
        """, (contor, an, luna, mzl_calculat, mzl_final,
               este_manual, indicator, zile_valide, zile_luna, now))

    def get_trafic_mzl(self, contor: str, an: int = None) -> pd.DataFrame:
        """Returnează DataFrame cu MZL final pentru un contor, opțional filtrat pe an."""
        params = [contor]
        where  = "contor = ?"
        if an is not None:
            where += " AND an = ?"; params.append(an)
        return pd.read_sql_query(
            f"SELECT * FROM trafic_mzl WHERE {where} ORDER BY an, luna",
            self._conn(), params=params
        )

    def get_all_trafic_mzl(self) -> pd.DataFrame:
        return pd.read_sql_query(
            "SELECT * FROM trafic_mzl ORDER BY contor, an, luna",
            self._conn()
        )

    # ── Statistici ────────────────────────────────────────────────────────────

    def stats(self) -> dict:
        conn = self._conn()
        n_rows     = conn.execute("SELECT COUNT(*) FROM inregistrari_orare").fetchone()[0]
        n_contoare = conn.execute("SELECT COUNT(DISTINCT contor) FROM inregistrari_orare").fetchone()[0]
        n_fisiere  = conn.execute("SELECT COUNT(*) FROM fisiere_procesate").fetchone()[0]
        size_mb    = os.path.getsize(self.db_path) / 1_048_576 if os.path.exists(self.db_path) else 0
        return {
            "randuri_orare": n_rows,
            "contoare":      n_contoare,
            "fisiere_log":   n_fisiere,
            "size_mb":       round(size_mb, 2),
        }

    def close(self):
        if hasattr(self._local, "conn") and self._local.conn:
            self._local.conn.close()
            self._local.conn = None


# =============================================================================
# CLASS ContoareDB
# =============================================================================

class ContoareDB:
    def __init__(self, db_path: str = CONTOARE_DB):
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self.db_path = db_path
        self._local  = threading.local()
        self._init_schema()

    def _conn(self) -> sqlite3.Connection:
        conn = getattr(self._local, "conn", None)
        if conn is not None:
            try:
                conn.execute("SELECT 1")
            except Exception:
                conn = None
        if conn is None:
            conn = sqlite3.connect(
                self.db_path,
                timeout=15,
                check_same_thread=False,
            )
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA journal_mode = WAL")
            conn.execute("PRAGMA busy_timeout = 10000")
            self._local.conn = conn
        return conn

    def _init_schema(self):
        conn = self._conn()
        conn.executescript(_CONTOARE_SCHEMA)
        # Migrare pentru tabele existente — adăugare coloane noi dacă lipsesc
        for _col, _tp in [("x","REAL"),("y","REAL"),("lat","REAL"),("lng","REAL"),("drdp","TEXT DEFAULT ''")]:
            try:
                conn.execute(f"ALTER TABLE contoare ADD COLUMN {_col} {_tp} DEFAULT NULL")
                conn.commit()
            except Exception:
                pass  # coloana există deja
        conn.commit()

    def _execute_with_retry(self, sql: str, params=(), retries: int = 5,
                             delay: float = 0.5) -> sqlite3.Cursor:
        """Retry automat la lock SQLite — pentru scrieri concurente multi-utilizator."""
        import time
        last_err = None
        for attempt in range(retries):
            try:
                conn = self._conn()
                cur = conn.execute(sql, params)
                conn.commit()
                return cur
            except sqlite3.OperationalError as e:
                last_err = e
                if "locked" in str(e).lower():
                    try:
                        self._local.conn.close()
                    except Exception:
                        pass
                    self._local.conn = None
                    time.sleep(delay * (attempt + 1))
                else:
                    raise
        raise sqlite3.OperationalError(
            f"DB locked dupa {retries} incercari: {last_err}")



    def get_all(self) -> dict:
        """Returnează dict {contor: {Drum, Pozitie_km, Localitate, Tip, IP, x, y, lat, lng}}"""
        rows = self._conn().execute(
            "SELECT * FROM contoare WHERE activ = 1 ORDER BY contor"
        ).fetchall()
        return {
            r["contor"]: {
                "Drum":       r["drum"],
                "Pozitie_km": r["pozitie_km"],
                "Localitate": r["localitate"],
                "Tip":        r["tip"],
                "IP":         r["ip"],
                "x":          r["x"],
                "y":          r["y"],
                "lat":        r["lat"],
                "lng":        r["lng"],
                "DRDP":       r["drdp"] if r["drdp"] is not None else "",
            }
            for r in rows
        }

    def get(self, contor: str) -> dict:
        row = self._conn().execute(
            "SELECT * FROM contoare WHERE contor = ?", (contor,)
        ).fetchone()
        if not row:
            return {}
        return {
            "Drum":       row["drum"],
            "Pozitie_km": row["pozitie_km"],
            "Localitate": row["localitate"],
            "Tip":        row["tip"],
            "IP":         row["ip"],
            "x":          row["x"],
            "y":          row["y"],
            "lat":        row["lat"],
            "lng":        row["lng"],
            "DRDP":       row["drdp"] if row["drdp"] is not None else "",
        }

    def upsert(self, contor: str, data: dict) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        existing = self._conn().execute(
            "SELECT contor FROM contoare WHERE contor = ?", (contor,)
        ).fetchone()
        if existing:
            self._execute_with_retry("""
                UPDATE contoare
                SET drum=?, pozitie_km=?, localitate=?, tip=?, ip=?,
                    x=?, y=?, lat=?, lng=?, drdp=?, modificat_la=?, activ=1
                WHERE contor=?
            """, (
                data.get("Drum", ""), data.get("Pozitie_km", ""),
                data.get("Localitate", ""), data.get("Tip", ""),
                data.get("IP", ""),
                data.get("x"), data.get("y"),
                data.get("lat"), data.get("lng"),
                data.get("DRDP", ""),
                now, contor
            ))
        else:
            self._execute_with_retry("""
                INSERT INTO contoare
                (contor, drum, pozitie_km, localitate, tip, ip,
                 x, y, lat, lng, drdp, activ, adaugat_la, modificat_la)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
            """, (
                contor, data.get("Drum", ""), data.get("Pozitie_km", ""),
                data.get("Localitate", ""), data.get("Tip", ""),
                data.get("IP", ""),
                data.get("x"), data.get("y"),
                data.get("lat"), data.get("lng"),
                data.get("DRDP", ""),
                now, now
            ))

    def save_all(self, db: dict) -> None:
        """Compatibilitate cu _save_contoare_db(db)."""
        for contor, data in db.items():
            self.upsert(contor, data)

    def delete(self, contor: str) -> None:
        """Soft delete cu retry."""
        self._execute_with_retry(
            "UPDATE contoare SET activ=0, modificat_la=? WHERE contor=?",
            (datetime.now().isoformat(timespec="seconds"), contor)
        )

    def get_localitate(self, contor: str) -> str:
        row = self._conn().execute(
            "SELECT localitate FROM contoare WHERE contor = ?", (contor,)
        ).fetchone()
        return row["localitate"] if row else ""

    def get_as_dataframe(self) -> pd.DataFrame:
        return pd.read_sql_query(
            "SELECT contor, drum, pozitie_km, localitate, tip, ip, drdp "
            "FROM contoare WHERE activ=1 ORDER BY contor",
            self._conn()
        )

    def import_from_dict(self, db: dict) -> int:
        count = 0
        for contor, data in db.items():
            self.upsert(contor, data)
            count += 1
        return count

    def close(self):
        if hasattr(self._local, "conn") and self._local.conn:
            self._local.conn.close()
            self._local.conn = None


# =============================================================================
# SINGLETONS
# =============================================================================

def get_traffic_db() -> TrafficDB:
    global _traffic_instance
    if _traffic_instance is None:
        with _lock:
            if _traffic_instance is None:
                _traffic_instance = TrafficDB()
    return _traffic_instance


def get_contoare_db() -> ContoareDB:
    global _contoare_instance
    if _contoare_instance is None:
        with _lock:
            if _contoare_instance is None:
                _contoare_instance = ContoareDB()
    return _contoare_instance


# =============================================================================
# MIGRARE — import date existente din JSON/Excel în SQLite
# =============================================================================

def migrate_contoare_from_json() -> int:
    """Migrează datele de identificare contoare din formatul vechi (Excel) în contoare.db."""
    from contoare_db import _load_contoare_db
    db_old = _load_contoare_db()
    if not db_old:
        print("[MIGRARE] Nicio dată de contor găsită în sursa veche.")
        return 0
    n = get_contoare_db().import_from_dict(db_old)
    print(f"[MIGRARE] {n} contoare migrate în contoare.db")
    return n


def migrate_traffic_from_excel(excel_folder: str,
                                tip_sursa: str = "PEEK") -> int:
    """Migrează date orare din Excel-urile Raport_Clase_* existente în trafic.db."""
    import glob
    pattern = "PEEK" if tip_sursa == "PEEK" else "VEK"
    files = glob.glob(
        os.path.join(excel_folder, f"Raport_Clase_{pattern}_*.xlsx"))

    total = 0
    tdb = get_traffic_db()
    for fp in files:
        try:
            df = pd.read_excel(fp, sheet_name="Date Detaliate")
            if df.empty:
                continue
            base    = os.path.splitext(os.path.basename(fp))[0]
            site_id = base.split("_")[-1]
            n = tdb.upsert_hourly_df(df, site_id,
                                     tip_sursa=tip_sursa,
                                     source_files=[fp])
            print(f"  [{site_id}] {n} rânduri")
            total += n
        except Exception as e:
            print(f"  Eroare {os.path.basename(fp)}: {e}")

    print(f"[MIGRARE] Total: {total} rânduri orare importate în trafic.db")
    return total
