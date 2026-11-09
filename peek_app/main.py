# =============================================================================
# main.py — Punct de intrare aplicație PEEK Traffic Analyzer
# =============================================================================
import sys
import os
import multiprocessing


def main():
    # Splash — apare INSTANT înainte de importurile grele
    from splash import SplashScreen
    splash = SplashScreen()
    splash.set_status("Se încarcă interfața...", 10)

    splash.set_status("Se încarcă modulele...", 30)
    import tkinter as tk

    splash.set_status("Se inițializează customtkinter...", 50)
    try:
        import customtkinter as ctk
        ctk_available = True
    except ImportError:
        ctk_available = False

    splash.set_status("Se încarcă procesoarele de date...", 70)
    from app import PeekApp

    splash.set_status("Gata!", 100)
    splash.close()

    if not ctk_available:
        print("[WARN] customtkinter nu este instalat.")
        print("       Rulează:  pip install customtkinter")
        print("       Aplicația pornește cu interfața de rezervă (tkinter standard).\n")

    app = PeekApp()
    app.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()   # PRIMA instrucțiune executată, obligatoriu pentru .exe
    main()