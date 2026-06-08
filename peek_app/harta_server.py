# =============================================================================
# harta_server.py — Server Flask minimal pentru harta contoare
# =============================================================================
# Pornit într-un thread daemon din PeekApp.__init__().
# Nu interferează cu GUI-ul Tkinter.
#
# Integrare în app.py:
#   1. Import la începutul fișierului (lângă celelalte importuri):
#        from harta_server import HartaServer
#
#   2. În PeekApp.__init__(), după ce GUI-ul e construit
#      (ex. după row_actions2 / btn_sorteaza):
#        self._harta_server = HartaServer()
#        self._harta_server.start()
#
#   3. Butonul (adaugă în row_actions2, după btn_sorteaza):
#        self.btn_harta = _ctk_btn(row_actions2, "🗺️  Hartă Contoare",
#                                  self._open_harta, "navy", width=200)
#        self.btn_harta.pack(side="left", padx=6)
#
#   4. Metodă în PeekApp:
#        def _open_harta(self):
#            self._harta_server.open_in_browser()
# =============================================================================

import threading
import webbrowser
import time
import os
import logging

from flask import Flask, jsonify, request, send_from_directory

from harta_api import _build_response

# ── Supress Flask/Werkzeug console spam ──────────────────────────────────────
log = logging.getLogger("werkzeug")
log.setLevel(logging.ERROR)

HARTA_PORT = 5757
HARTA_URL  = f"http://127.0.0.1:{HARTA_PORT}"

# ── Folder unde e harta_contoare.html ────────────────────────────────────────
# În executabilul PyInstaller fișierele din datas sunt extrase în sys._MEIPASS
import sys as _sys
_HERE = getattr(_sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


class HartaServer:
    """
    Server Flask rulând în thread daemon.
    Sigur de pornit înainte ca GUI-ul să fie complet inițializat.
    """

    def __init__(self):
        self._started  = False
        self._flask    = Flask(__name__)
        self._register_routes()

    def _register_routes(self):
        flask_app = self._flask

        @flask_app.route("/")
        def index():
            return send_from_directory(_HERE, "harta_contoare.html")

        @flask_app.route("/api/harta_contoare", methods=["GET"])
        def api_contoare():
            try:
                return jsonify(_build_response())
            except Exception as e:
                return jsonify({"error": str(e)}), 500

        @flask_app.route("/api/harta_contoare/coordonate", methods=["POST"])
        def api_coordonate():
            try:
                from database import get_contoare_db
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
                return jsonify({"ok": True, "contor": contor,
                                "lat": lat, "lng": lng})
            except Exception as e:
                return jsonify({"error": str(e)}), 500

    def start(self):
        """Pornește serverul Flask în thread daemon (non-blocking)."""
        if self._started:
            return
        self._started = True

        def _run():
            self._flask.run(
                host="127.0.0.1",
                port=HARTA_PORT,
                debug=False,
                use_reloader=False,
                threaded=True,
            )

        t = threading.Thread(target=_run, daemon=True, name="HartaFlask")
        t.start()
        # Mică pauză ca serverul să fie gata înainte de primul open_in_browser
        time.sleep(0.8)

    def open_in_browser(self):
        """Deschide harta în browser-ul default."""
        if not self._started:
            self.start()
        webbrowser.open(HARTA_URL)
