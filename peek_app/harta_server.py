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
import re
import socket
import subprocess
import tempfile
import logging

from flask import Flask, jsonify, request, send_from_directory

from harta_api import _build_response, reconstruieste_cache_async

# ── Supress Flask/Werkzeug console spam ──────────────────────────────────────
log = logging.getLogger("werkzeug")
log.setLevel(logging.ERROR)

HARTA_PORT = 5757
HARTA_URL  = f"http://127.0.0.1:{HARTA_PORT}"

# ── Conectare Viper prin Remote Desktop ──────────────────────────────────────
# Parola NU se salvează nicăieri: Windows o cere la fiecare conectare.
VIPER_HOST      = "192.168.104.157"
VIPER_USER      = "user"
VIPER_RDP_PORT  = 3389
_ORIGINI_OK     = {HARTA_URL, f"http://localhost:{HARTA_PORT}"}
_CONTOR_RE      = re.compile(r"^[A-Za-z0-9_.\- ]{1,40}$")

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
                return jsonify(_build_response(
                    force_refresh=request.args.get("refresh") == "1"))
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

        @flask_app.route("/api/viper/conecteaza", methods=["POST"])
        def api_viper():
            """
            Pornește Remote Desktop către VM-ul cu Viper și pune codul contorului
            în clipboard (partajat în sesiunea RDP). Parola o cere Windows.
            """
            try:
                # Doar din pagina hărții (blochează apeluri din alte site-uri)
                if request.headers.get("Origin") not in _ORIGINI_OK:
                    return jsonify({"error": "origine nepermisă"}), 403
                if os.name != "nt":
                    return jsonify({"error": "disponibil doar pe Windows"}), 400

                data   = request.get_json(force=True) or {}
                contor = str(data.get("contor", "")).strip()
                if not _CONTOR_RE.match(contor):
                    return jsonify({"error": "cod contor invalid"}), 400

                # 1) VM-ul răspunde? (altfel VPN-ul e probabil oprit)
                try:
                    socket.create_connection(
                        (VIPER_HOST, VIPER_RDP_PORT), timeout=3).close()
                except OSError:
                    return jsonify({
                        "error": f"VM-ul Viper ({VIPER_HOST}) nu răspunde. "
                                 f"Verifică conexiunea VPN."}), 503

                # 2) Date contor (SiteID Viper = cod cu zerouri în față; IP din DB)
                ip = ""
                try:
                    from database import get_contoare_db
                    ip = (get_contoare_db().get(contor) or {}).get("IP", "") or ""
                except Exception:
                    pass
                site_id = contor.zfill(12) if contor.isdigit() else contor

                # 3) Clipboard (se transferă în sesiunea RDP)
                try:
                    subprocess.run(["clip"], input=site_id.encode("ascii", "ignore"),
                                   creationflags=0x08000000, timeout=5)
                except Exception:
                    pass

                # 4) Fișier .rdp (fără parolă; Windows o cere de fiecare dată)
                rdp_path = os.path.join(tempfile.gettempdir(), "viper_cestrin.rdp")
                with open(rdp_path, "w", encoding="utf-8") as f:
                    f.write(
                        f"full address:s:{VIPER_HOST}:{VIPER_RDP_PORT}\n"
                        f"username:s:{VIPER_USER}\n"
                        "prompt for credentials:i:1\n"
                        "redirectclipboard:i:1\n"
                        "enablecredsspsupport:i:1\n"
                        "authentication level:i:2\n"
                    )
                subprocess.Popen(["mstsc", rdp_path])

                return jsonify({"ok": True, "contor": contor,
                                "site_id": site_id, "ip": ip})
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
