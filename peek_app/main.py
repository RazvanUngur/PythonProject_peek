# =============================================================================
# main.py — Punct de intrare aplicație PEEK Traffic Analyzer
# =============================================================================

import sys
import os

# ── Splash screen — apare INSTANT înainte de importurile grele ────────────────
from splash import SplashScreen
splash = SplashScreen()
splash.set_status("Se încarcă interfața...", 10)

# ── Importuri grele — pandas, openpyxl, customtkinter ─────────────────────────
splash.set_status("Se încarcă modulele...", 30)
import tkinter as tk

splash.set_status("Se inițializează customtkinter...", 50)
try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False

splash.set_status("Se încarcă procesoarele de date...", 70)
from app import PeekApp

splash.set_status("Gata!", 100)
splash.close()

# ── Pornire aplicație ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    if not CTK_AVAILABLE:
        print("[WARN] customtkinter nu este instalat.")
        print("       Rulează:  pip install customtkinter")
        print("       Aplicația pornește cu interfața de rezervă (tkinter standard).\n")
    app = PeekApp()
    app.mainloop()
