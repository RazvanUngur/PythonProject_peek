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

import sqlite3
import calendar
from datetime import datetime
from flask import jsonify, request

from database import (get_traffic_db, get_contoare_db,
                      TRAFFIC_DB, CONTOARE_DB)


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


# ── Statistici per contor pentru luna de referință ───────────────────────────
def _stats_contor(contor: str, an_ref: int, luna_ref: int) -> dict:
    """
    Calculează statusul unui contor pentru luna de referință.

    MZL final:
      1. Dacă există intrare în mzl_manual → acea valoare
      2. Altfel calculat din inregistrari_orare:
         total_general / zile_calendaristice (dacă ≥20 zile) sau / zile_cu_date

    Status (folosit de frontend pentru culoare marker):
      CLASIFICATOR  — are date luna ref, clasa_15 < 10% din total
      TOTALIZATOR   — are date luna ref, clasa_15 ≥ 10% din total
      FARA_COM      — nu are date luna ref, dar ultima lună înregistrată are total > 0
      DEFECT        — luna ref: total = 0, SAU nu are date luna ref și ultima lună total = 0
    """
    conn = sqlite3.connect(TRAFFIC_DB, timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        # ── 1. Date pentru luna de referință ──────────────────────────────────
        row_ref = conn.execute("""
            SELECT
                SUM(total_general)                           AS total_veh,
                SUM(b1_clasa_15 + b2_clasa_15 + b3_clasa_15 +
                    b4_clasa_15 + b5_clasa_15 + b6_clasa_15) AS total_cls15,
                COUNT(*)                                     AS ore_inreg,
                COUNT(DISTINCT zi)                           AS zile_cu_date
            FROM inregistrari_orare
            WHERE contor = ? AND an = ? AND luna = ?
        """, (contor, an_ref, luna_ref)).fetchone()

        are_date_ref = (row_ref and (row_ref["ore_inreg"] or 0) > 0)

        # ── 2. MZL final pentru luna de referință ─────────────────────────────
        mzl_final = None
        mzl_sursa = None  # "trafic_mzl" | "manual" | "calculat"

        # 2a. trafic_mzl — sursa principală (calculată de aplicație la procesare)
        tmzl_row = conn.execute(
            "SELECT mzl_final, este_manual FROM trafic_mzl "
            "WHERE contor=? AND an=? AND luna=?",
            (contor, an_ref, luna_ref)
        ).fetchone()

        if tmzl_row:
            mzl_final = tmzl_row["mzl_final"]
            mzl_sursa = "manual" if tmzl_row["este_manual"] else "trafic_mzl"
        else:
            # 2b. Fallback: mzl_manual (dacă nu a fost procesat prin raport)
            mzl_row = conn.execute(
                "SELECT mzl_valoare FROM mzl_manual "
                "WHERE contor=? AND an=? AND luna=?",
                (contor, an_ref, luna_ref)
            ).fetchone()

            if mzl_row:
                mzl_final = mzl_row["mzl_valoare"]
                mzl_sursa = "manual"
            elif are_date_ref:
                # 2c. Fallback: calcul direct din orare
                veh  = row_ref["total_veh"] or 0
                zile = row_ref["zile_cu_date"] or 1
                zile_cal = calendar.monthrange(an_ref, luna_ref)[1]
                zile_div = zile_cal if zile >= 20 else zile
                mzl_final = round(veh / zile_div) if zile_div > 0 else 0
                mzl_sursa = "calculat"

        # ── 3. Ultima lună cu date (pentru FARA_COM / DEFECT când ref lipsește)
        row_ultima = conn.execute("""
            SELECT an, luna, SUM(total_general) AS total_veh
            FROM inregistrari_orare
            WHERE contor = ?
            GROUP BY an, luna
            ORDER BY an DESC, luna DESC
            LIMIT 1
        """, (contor,)).fetchone()

        total_ultima = int(row_ultima["total_veh"] or 0) if row_ultima else 0
        ultima_luna_str = (
            f"{row_ultima['luna']:02d}.{row_ultima['an']}"
            if row_ultima else None
        )

        # ── 4. Status ──────────────────────────────────────────────────────────
        if are_date_ref:
            total_veh  = int(row_ref["total_veh"]   or 0)
            total_cls15= int(row_ref["total_cls15"] or 0)
            if total_veh == 0:
                status = "DEFECT"
            elif total_cls15 / total_veh >= 0.10:
                status = "TOTALIZATOR"
            else:
                status = "CLASIFICATOR"
        else:
            # Nu are date pe luna de referință
            if total_ultima > 0:
                status = "FARA_COM"
            else:
                status = "DEFECT"

        # ── 5. Acoperire ore (informativă) ────────────────────────────────────
        acop_pct = 0.0
        if are_date_ref:
            zile_cal  = calendar.monthrange(an_ref, luna_ref)[1]
            ore_inreg = int(row_ref["ore_inreg"] or 0)
            acop_pct  = round(ore_inreg / (zile_cal * 24) * 100, 1)

        return {
            "mzl":          mzl_final,
            "mzl_sursa":    mzl_sursa,  # "trafic_mzl" | "manual" | "calculat" | None
            "luna_ref":     f"{luna_ref:02d}.{an_ref}",
            "status":       status,
            "are_date_ref": are_date_ref,
            "acop_pct":     acop_pct,
            "ultima_luna":  ultima_luna_str,
            "total_cls15":  int(row_ref["total_cls15"] or 0) if are_date_ref else None,
            "total_veh":    int(row_ref["total_veh"]   or 0) if are_date_ref else None,
        }

    finally:
        conn.close()


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
def _build_response() -> list:
    an_ref, luna_ref = _luna_referinta()

    cdb = get_contoare_db()
    tdb = get_traffic_db()

    contoare_dict    = cdb.get_all()
    contoare_cu_date = set(tdb.get_contoare_disponibile())
    coduri_vechi     = _get_coduri_vechi()  # exclude codurile redenumite

    # Excludem codurile vechi din ambele surse
    toti = sorted(
        (set(contoare_dict.keys()) | contoare_cu_date) - coduri_vechi
    )

    result = []
    for contor in toti:
        info  = contoare_dict.get(contor, {})
        stats = _stats_contor(contor, an_ref, luna_ref)

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

    return result


# ── Flask routes ──────────────────────────────────────────────────────────────
def register_harta_routes(app):

    @app.route("/api/harta_contoare", methods=["GET"])
    def api_harta_contoare():
        try:
            return jsonify(_build_response())
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
