# =============================================================================
# harta_api.py — Endpoint Flask pentru Harta Contoare
# =============================================================================
#
# Adaugă în app.py:
#   from harta_api import register_harta_routes
#   register_harta_routes(app)
#
# Endpoint: GET  /api/harta_contoare
# Endpoint: POST /api/harta_contoare/coordonate  { contor, lat, lng }
# =============================================================================

import os
import json
import sqlite3
import calendar
import tempfile
import threading
from datetime import datetime
from flask import jsonify, request

from database import (get_traffic_db, get_contoare_db,
                      TRAFFIC_DB, CONTOARE_DB)


# ── Cache pe disc ─────────────────────────────────────────────────────────────
# Statisticile grele (MZL, status, acoperire) se calculează O DATĂ, după
# procesarea fișierelor, și se salvează aici. Harta doar citește fișierul.
# Metadatele (coordonate, drum, activ/inactiv) se citesc mereu live din
# contoare.db — deci mutarea unui marker sau debifarea unui contor
# apar imediat, fără rebuild.
CACHE_FILE  = os.path.join(os.path.dirname(os.path.abspath(TRAFFIC_DB)),
                           "harta_cache.json")
_cache_lock = threading.Lock()

# Versiunea formatului cache-ului. Crește când se schimbă structura din "stats"
# (2 = adaugă detaliul pe benzi). Un cache cu altă versiune e ignorat și se
# reconstruiește singur la prima deschidere a hărții.
CACHE_VERSIUNE = 2

# ── Benzi ─────────────────────────────────────────────────────────────────────
PRAG_TOT_BANDA = 0.10     # bandă totalizator dacă clasa 15 / total bandă >= prag
MAX_BENZI      = 6
PERECHI        = {2: [(1, 2)], 4: [(1, 4), (2, 3)]}   # ca în raport_mod_functionare
RAPORT_PERECHE = 2        # banda mai mică are de >= 2 ori mai puțin trafic decât perechea


# ── Luna de referință ─────────────────────────────────────────────────────────
def _luna_referinta() -> tuple:
    """
    Returnează (an, luna) de referință pentru harta:
      - după ziua 10 a lunii curente → luna curentă - 1
      - înainte sau pe ziua 10       → luna curentă - 2
    """
    azi = datetime.now()
    decalaj = 1 if azi.day > 10 else 2
    luna = azi.month - decalaj
    an   = azi.year
    while luna <= 0:
        luna += 12
        an   -= 1
    return an, luna


# ── Interogări (folosesc indexurile existente din database.py) ───────────────
def _cols_benzi() -> str:
    """Sume lunare pe fiecare bandă: total, clasa 15, clasa 2, clasa 4."""
    parts = ["MAX(n_benzi) AS n_benzi"]
    for b in range(1, MAX_BENZI + 1):
        parts.append(f"COALESCE(SUM(total_b{b}), 0)     AS t{b}")
        parts.append(f"COALESCE(SUM(b{b}_clasa_15), 0)  AS c{b}")
        parts.append(f"COALESCE(SUM(b{b}_clasa_2), 0)   AS k2_{b}")
        parts.append(f"COALESCE(SUM(b{b}_clasa_4), 0)   AS k4_{b}")
    return ",\n        ".join(parts)


_SQL_REF = """
    SELECT
        contor,
        SUM(total_general)                           AS total_veh,
        SUM(b1_clasa_15 + b2_clasa_15 + b3_clasa_15 +
            b4_clasa_15 + b5_clasa_15 + b6_clasa_15) AS total_cls15,
        COUNT(*)                                     AS ore_inreg,
        COUNT(DISTINCT zi)                           AS zile_cu_date,
        """ + _cols_benzi() + """
    FROM inregistrari_orare
    WHERE an = ? AND luna = ? {extra}
    GROUP BY contor
"""

_SQL_ULTIMA = """
    SELECT an, luna, SUM(total_general) AS total_veh
    FROM inregistrari_orare
    WHERE contor = ?
    GROUP BY an, luna
    ORDER BY an DESC, luna DESC
    LIMIT 1
"""


# ── Starea pe benzi ───────────────────────────────────────────────────────────
def _analiza_benzi(row, total_general: int) -> list:
    """
    Starea fiecărei benzi pe luna de referință (sume lunare):
      D — total bandă == 0                       (nu înregistrează)
      T — clasa 15 / total >= PRAG_TOT_BANDA     (senzor piezoelectric)
          SAU clasa 2 < clasa 4                  (buclă inductivă de ieșire)
      C — altfel                                 (clasificator)
    `incomplet` = banda mai mică dintr-o pereche (2 sau 4 benzi) cu trafic > 0,
    când perechea ei are >= RAPORT_PERECHE x mai mult. Doar informativ.
    Returnează [] dacă nu există detaliu pe benzi în baza de date.
    """
    try:
        n = int(row["n_benzi"] or 2)
        n = n if n in (2, 4, 6) else max(2, min(n, MAX_BENZI))
        tot = {b: int(row[f"t{b}"] or 0)    for b in range(1, n + 1)}
        c15 = {b: int(row[f"c{b}"] or 0)    for b in range(1, n + 1)}
        c2  = {b: int(row[f"k2_{b}"] or 0)  for b in range(1, n + 1)}
        c4  = {b: int(row[f"k4_{b}"] or 0)  for b in range(1, n + 1)}
    except (KeyError, IndexError, TypeError, ValueError):
        return []

    # Total general > 0 dar nicio bandă cu trafic => înregistrări vechi fără
    # detaliu pe benzi; nu le tratăm ca defecte.
    if total_general > 0 and not any(tot.values()):
        return []

    incomplete = set()
    for a, b in PERECHI.get(n, []):
        mic, mare = (a, b) if tot[a] < tot[b] else (b, a)
        if tot[mic] > 0 and tot[mare] >= RAPORT_PERECHE * tot[mic]:
            incomplete.add(mic)

    benzi = []
    for b in range(1, n + 1):
        t = tot[b]
        if t == 0:
            stare, piezo, k = "D", False, False
        else:
            piezo = c15[b] / t >= PRAG_TOT_BANDA
            k     = c2[b] < c4[b]
            stare = "T" if (piezo or k) else "C"
        benzi.append({
            "b":            b,
            "stare":        stare,
            "total":        t,
            "pct15":        round(c15[b] / t * 100, 1) if t else 0.0,
            "piezo":        piezo,
            "cls2_lt_cls4": k,
            "incomplet":    b in incomplete,
        })
    return benzi


def _status_din_benzi(benzi: list) -> str:
    """
    Statusul contorului din starea benzilor (în ordinea priorității):
      DEFECT              — nicio bandă nu înregistrează
      PARTIAL_DEFECT      — cel puțin o bandă nu înregistrează, dar există
                            minim o bandă clasificator/totalizator
      PARTIAL_TOTALIZATOR — cel puțin o bandă totalizator și minim o bandă
                            clasificator, SAU cel puțin o bandă cu clasa 2 < clasa 4
      TOTALIZATOR         — toate benzile active sunt totalizator
      CLASIFICATOR        — toate benzile active sunt clasificator
    """
    active = [x for x in benzi if x["stare"] != "D"]
    if not active:
        return "DEFECT"
    if len(active) < len(benzi):
        return "PARTIAL_DEFECT"
    n_tot = sum(1 for x in active if x["stare"] == "T")
    n_cls = len(active) - n_tot
    if any(x["cls2_lt_cls4"] for x in active) or (n_tot and n_cls):
        return "PARTIAL_TOTALIZATOR"
    return "TOTALIZATOR" if n_tot else "CLASIFICATOR"


# ── Logica de status (comună pentru calcul individual și în masă) ────────────
def _compune_stats(an_ref, luna_ref, row_ref, tmzl_row, mzl_row, row_ultima):
    """
    MZL final:
      1. trafic_mzl (calculat la procesare; manual dacă este_manual)
      2. mzl_manual
      3. calcul direct din inregistrari_orare

    Status (cu date pe luna ref, calculat din starea benzilor — vezi
    _status_din_benzi): CLASIFICATOR, TOTALIZATOR, PARTIAL_TOTALIZATOR,
    PARTIAL_DEFECT, DEFECT. Dacă nu există detaliu pe benzi, se folosește
    vechea regulă pe total: clasa_15 ≥ 10% → TOTALIZATOR, altfel CLASIFICATOR.
      FARA_COM      — nu are date luna ref, dar ultima lună înregistrată are total > 0
      DEFECT        — luna ref: total = 0, SAU nu are date luna ref și ultima lună total = 0
    """
    are_date_ref = bool(row_ref and (row_ref["ore_inreg"] or 0) > 0)

    mzl_final = None
    mzl_sursa = None
    if tmzl_row:
        mzl_final = tmzl_row["mzl_final"]
        mzl_sursa = "manual" if tmzl_row["este_manual"] else "trafic_mzl"
    elif mzl_row:
        mzl_final = mzl_row["mzl_valoare"]
        mzl_sursa = "manual"
    elif are_date_ref:
        veh  = row_ref["total_veh"] or 0
        zile = row_ref["zile_cu_date"] or 1
        zile_cal = calendar.monthrange(an_ref, luna_ref)[1]
        zile_div = zile_cal if zile >= 20 else zile
        mzl_final = round(veh / zile_div) if zile_div > 0 else 0
        mzl_sursa = "calculat"

    total_ultima = int(row_ultima["total_veh"] or 0) if row_ultima else 0
    ultima_luna_str = (f"{row_ultima['luna']:02d}.{row_ultima['an']}"
                       if row_ultima else None)

    benzi = []
    if are_date_ref:
        total_veh   = int(row_ref["total_veh"]   or 0)
        total_cls15 = int(row_ref["total_cls15"] or 0)
        benzi = _analiza_benzi(row_ref, total_veh)
        if total_veh == 0:
            status = "DEFECT"
        elif benzi:
            status = _status_din_benzi(benzi)
        elif total_cls15 / total_veh >= 0.10:
            status = "TOTALIZATOR"
        else:
            status = "CLASIFICATOR"
    else:
        status = "FARA_COM" if total_ultima > 0 else "DEFECT"

    acop_pct = 0.0
    if are_date_ref:
        zile_cal = calendar.monthrange(an_ref, luna_ref)[1]
        acop_pct = round(int(row_ref["ore_inreg"] or 0) / (zile_cal * 24) * 100, 1)

    return {
        "mzl":          mzl_final,
        "mzl_sursa":    mzl_sursa,
        "luna_ref":     f"{luna_ref:02d}.{an_ref}",
        "status":       status,
        "are_date_ref": are_date_ref,
        "acop_pct":     acop_pct,
        "ultima_luna":  ultima_luna_str,
        "total_cls15":  int(row_ref["total_cls15"] or 0) if are_date_ref else None,
        "total_veh":    int(row_ref["total_veh"]   or 0) if are_date_ref else None,
        "benzi":        benzi,
    }


# ── Un singur contor (folosit doar ca fallback pentru contoare fără cache) ───
def _stats_contor(contor: str, an_ref: int, luna_ref: int) -> dict:
    conn = sqlite3.connect(TRAFFIC_DB, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        row_ref = conn.execute(
            _SQL_REF.format(extra="AND contor = ?"),
            (an_ref, luna_ref, contor)).fetchone()
        tmzl = conn.execute(
            "SELECT mzl_final, este_manual FROM trafic_mzl "
            "WHERE contor=? AND an=? AND luna=?",
            (contor, an_ref, luna_ref)).fetchone()
        mzl = conn.execute(
            "SELECT mzl_valoare FROM mzl_manual "
            "WHERE contor=? AND an=? AND luna=?",
            (contor, an_ref, luna_ref)).fetchone()
        ultima = conn.execute(_SQL_ULTIMA, (contor,)).fetchone()
        return _compune_stats(an_ref, luna_ref, row_ref, tmzl, mzl, ultima)
    finally:
        conn.close()


# ── TOATE contoarele dintr-un foc (câteva interogări, nu 4 × N) ──────────────
def _conn_ro():
    """Conexiune doar-citire. Dacă baza (WAL, pe rețea) nu permite mode=ro,
    cădem pe conexiunea normală — folosită EXCLUSIV pentru SELECT."""
    try:
        uri = "file:" + os.path.abspath(TRAFFIC_DB).replace("\\", "/") + "?mode=ro"
        conn = sqlite3.connect(uri, uri=True, timeout=60)
        conn.execute("SELECT 1 FROM sqlite_master LIMIT 1")
    except sqlite3.OperationalError:
        conn = sqlite3.connect(TRAFFIC_DB, timeout=60)
    conn.row_factory = sqlite3.Row
    return conn


def _contoare_cu_date() -> set:
    """
    Echivalent rapid cu `SELECT DISTINCT contor FROM inregistrari_orare`.
    DISTINCT parcurge tot indexul (milioane de rânduri); aici sărim de la un
    contor la următorul cu câte o căutare în index (~1 căutare / contor).
    """
    conn = _conn_ro()
    try:
        rows = conn.execute("""
            WITH RECURSIVE c(contor) AS (
                SELECT MIN(contor) FROM inregistrari_orare
                UNION ALL
                SELECT (SELECT MIN(contor) FROM inregistrari_orare
                        WHERE contor > c.contor)
                FROM c WHERE c.contor IS NOT NULL
            )
            SELECT contor FROM c WHERE contor IS NOT NULL
        """).fetchall()
        return {r[0] for r in rows}
    finally:
        conn.close()


def _stats_unul(conn, contor, an_ref, luna_ref, tmzl, mzl):
    """Statistici pentru UN contor, doar interogări filtrate pe contor
    (folosesc indexul existent; nu scanează tot tabelul)."""
    row_ref = conn.execute(
        _SQL_REF.format(extra="AND contor = ?"),
        (an_ref, luna_ref, contor)).fetchone()
    are_ref = bool(row_ref and (row_ref["ore_inreg"] or 0) > 0)

    # ultima lună = un singur seek, nu GROUP BY pe tot istoricul contorului
    u = conn.execute(
        "SELECT an, luna FROM inregistrari_orare WHERE contor = ? "
        "ORDER BY an DESC, luna DESC LIMIT 1", (contor,)).fetchone()
    ultima = None
    if u:
        total = None
        if not are_ref:   # totalul ultimei luni contează doar fără date pe luna ref
            t = conn.execute(
                "SELECT SUM(total_general) FROM inregistrari_orare "
                "WHERE contor = ? AND an = ? AND luna = ?",
                (contor, u["an"], u["luna"])).fetchone()
            total = t[0] if t else 0
        ultima = {"an": u["an"], "luna": u["luna"], "total_veh": total or 0}

    return _compune_stats(an_ref, luna_ref, row_ref if are_ref else None,
                          tmzl.get(contor), mzl.get(contor), ultima)


def _stats_lista(contoare, an_ref, luna_ref, log=None, workers=4) -> dict:
    """Calculează statisticile pentru o listă de contoare, în paralel."""
    import time as _t
    from concurrent.futures import ThreadPoolExecutor
    contoare = sorted({str(c).strip() for c in contoare})
    n = len(contoare)
    t0 = _t.time()

    def _p(msg):
        print("[harta] " + msg)
        if log:
            try: log("   " + msg)
            except Exception: pass

    conn = _conn_ro()
    try:
        tmzl = {r["contor"]: r for r in conn.execute(
            "SELECT contor, mzl_final, este_manual FROM trafic_mzl "
            "WHERE an=? AND luna=?", (an_ref, luna_ref))}
        mzl = {r["contor"]: r for r in conn.execute(
            "SELECT contor, mzl_valoare FROM mzl_manual "
            "WHERE an=? AND luna=?", (an_ref, luna_ref))}
    finally:
        conn.close()
    _p(f"MZL încărcat; calculez {n} contoare ({workers} fire paralele)...")

    rezultat, done = {}, [0]
    lock = threading.Lock()
    local = threading.local()

    def _lucru(c):
        if not hasattr(local, "conn"):
            local.conn = _conn_ro()
        r = _stats_unul(local.conn, c, an_ref, luna_ref, tmzl, mzl)
        with lock:
            rezultat[c] = r
            done[0] += 1
            if done[0] % 10 == 0 or done[0] == n:
                _p(f"  {done[0]}/{n} contoare  ({_t.time()-t0:.0f}s)")
        return c

    with ThreadPoolExecutor(max_workers=workers) as ex:
        list(ex.map(_lucru, contoare))
    return rezultat


# ── Cache: citire / scriere / reconstruire ───────────────────────────────────
def _citeste_cache(an_ref: int, luna_ref: int):
    """Returnează conținutul cache-ului (dict cu 'stats' și 'disponibile')
    dacă e valid pentru luna de referință curentă, altfel None."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if (data.get("versiune") == CACHE_VERSIUNE
                and data.get("luna_ref") == f"{luna_ref:02d}.{an_ref}"):
            if data.get("stats"):
                return data
    except Exception:
        pass
    return None


def _scrie_cache(an_ref: int, luna_ref: int, stats: dict, disponibile=None):
    data = {
        "versiune":   CACHE_VERSIUNE,
        "disponibile": sorted(disponibile) if disponibile is not None else None,
        "luna_ref":   f"{luna_ref:02d}.{an_ref}",
        "generat_la": datetime.now().isoformat(timespec="seconds"),
        "stats":      stats,
    }
    d = os.path.dirname(CACHE_FILE)
    fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, CACHE_FILE)        # atomic — nimeni nu citește fișier pe jumătate
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def reconstruieste_cache(log=None, contoare=None) -> dict:
    """
    contoare=None  → reconstruire COMPLETĂ (toate contoarele).
    contoare=[...] → actualizare INCREMENTALĂ: doar acele contoare (ex. cele
                     tocmai procesate), peste cache-ul existent. Dacă nu există
                     cache valid pentru luna curentă, se face reconstruire completă.
    """
    import time as _t
    def _p(msg):
        print("[harta] " + msg)
        if log:
            try: log("   " + msg)
            except Exception: pass

    with _cache_lock:
        t0 = _t.time()
        an_ref, luna_ref = _luna_referinta()
        vechi = _citeste_cache(an_ref, luna_ref) if contoare else None

        if vechi:                                   # ── incremental
            _p(f"actualizare incrementală: {len(set(contoare))} contoare")
            noi = _stats_lista(contoare, an_ref, luna_ref, log=log)
            stats = dict(vechi["stats"])
            stats.update(noi)
            disponibile = sorted(set(vechi.get("disponibile") or stats.keys())
                                 | set(noi.keys()))
        else:                                       # ── complet
            _p("reconstruire completă: citesc lista contoarelor...")
            t = _t.time()
            try:
                lista = _contoare_cu_date()
            except Exception:
                try:   # fallback: metoda veche din database.py (lentă)
                    lista = set(get_traffic_db().get_contoare_disponibile())
                except Exception:
                    lista = set()
            _p(f"lista contoarelor: {len(lista)} în {_t.time()-t:.0f}s")
            stats = _stats_lista(lista, an_ref, luna_ref, log=log)
            disponibile = sorted(lista)

        try:
            _scrie_cache(an_ref, luna_ref, stats, disponibile)
        except Exception as e:
            _p(f"⚠ cache nescris: {e}")
        _p(f"cache gata în {_t.time()-t0:.0f}s")
        return {"stats": stats, "disponibile": disponibile}


def reconstruieste_cache_async(contoare=None):
    """Variantă non-blocking, sigură de apelat din GUI."""
    threading.Thread(target=reconstruieste_cache, kwargs={"contoare": contoare},
                     daemon=True, name="HartaCacheRebuild").start()


# ── Coduri vechi redenumite (de exclus din hartă) ────────────────────────────
def _get_coduri_vechi() -> set:
    """
    Returnează setul de coduri vechi din contor_alias — acestea nu mai
    trebuie afișate pe hartă, datele lor au fost migrate la codul nou.
    """
    try:
        conn = sqlite3.connect(CONTOARE_DB, timeout=10)
        try:
            rows = conn.execute("SELECT cod_vechi FROM contor_alias").fetchall()
            return {r[0] for r in rows}
        finally:
            conn.close()
    except Exception:
        return set()  # tabelul nu există încă — nicio excludere


# ── Build răspuns complet ─────────────────────────────────────────────────────
def _build_response(force_refresh: bool = False) -> list:
    import time as _t
    _t0 = _t.time()
    an_ref, luna_ref = _luna_referinta()

    # Statistici grele + lista contoarelor: din cache (instant) sau reconstruite
    cache = None if force_refresh else _citeste_cache(an_ref, luna_ref)
    din_cache = cache is not None
    if cache is None:
        cache = reconstruieste_cache()
    stats_map = cache["stats"]
    _t1 = _t.time()

    cdb = get_contoare_db()
    contoare_dict = cdb.get_all()          # doar contoarele ACTIVE
    if cache.get("disponibile") is not None:
        contoare_cu_date = set(cache["disponibile"])
    else:
        contoare_cu_date = set(get_traffic_db().get_contoare_disponibile())
    coduri_vechi = _get_coduri_vechi()     # exclude codurile redenumite
    inactivi     = cdb.get_inactive()      # debifate în Gestionare Contoare
    _t2 = _t.time()

    # Excludem codurile vechi și contoarele inactive din ambele surse
    toti = sorted(
        (set(contoare_dict.keys()) | contoare_cu_date) - coduri_vechi - inactivi
    )

    result = []
    for contor in toti:
        info  = contoare_dict.get(contor, {})
        # Contor absent din cache = nu are nicio înregistrare în DB de trafic
        # (nici orare, nici MZL) -> status calculat direct, fără interogări.
        stats = dict(stats_map.get(contor) or
                     _compune_stats(an_ref, luna_ref, None, None, None, None))

        # Contoare fără coordonate → status special
        lat = info.get("lat")
        lng = info.get("lng")
        if not lat or not lng:
            stats["status"] = "NELOCALIZAT"

        result.append({
            "contor":       contor,
            "drum":         info.get("Drum", ""),
            "pozitie_km":   info.get("Pozitie_km", ""),
            "localitate":   info.get("Localitate", ""),
            "tip":          info.get("Tip", ""),
            "ip":           info.get("IP", ""),
            "drdp":         info.get("DRDP", ""),
            "lat":          lat,
            "lng":          lng,
            **stats,
        })

    print(f"[harta] cache={'DA' if din_cache else 'reconstruit'} "
          f"stats={_t1-_t0:.1f}s  contoare.db={_t2-_t1:.1f}s  "
          f"total={_t.time()-_t0:.1f}s  ({len(result)} contoare)")
    return result


# ── Flask routes ──────────────────────────────────────────────────────────────
def register_harta_routes(app):

    @app.route("/api/harta_contoare", methods=["GET"])
    def api_harta_contoare():
        try:
            return jsonify(_build_response(
                force_refresh=request.args.get("refresh") == "1"))
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/api/harta_contoare/coordonate", methods=["POST"])
    def api_salveaza_coordonate():
        """
        Salvează coordonatele unui contor mutat manual pe hartă.
        Body JSON: { "contor": "DN1_001", "lat": 45.123, "lng": 25.456 }
        """
        try:
            data   = request.get_json(force=True)
            contor = data.get("contor", "").strip()
            lat    = float(data["lat"])
            lng    = float(data["lng"])

            if not contor:
                return jsonify({"error": "contor lipsă"}), 400
            if not (-90 <= lat <= 90) or not (-180 <= lng <= 180):
                return jsonify({"error": "coordonate invalide"}), 400

            cdb  = get_contoare_db()
            info = cdb.get(contor) or {}
            info["lat"] = lat
            info["lng"] = lng
            cdb.upsert(contor, info)

            return jsonify({"ok": True, "contor": contor, "lat": lat, "lng": lng})
        except Exception as e:
            return jsonify({"error": str(e)}), 500
