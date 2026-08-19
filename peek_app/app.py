# =============================================================================
# app.py — Fereastra principală PeekApp
# =============================================================================

import os
import re
import sys
import time
import threading
from datetime import datetime
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    ctk = None
    CTK_AVAILABLE = False

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from config import (
    CTK_THEME, CTK_COLOR,
    CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME, RAPOARTE_PEEK_FOLDER,
    FONT_SMALL, FONT_MONO,
    BG_APP, BG_CARD, BG_LOG, BG_TITLE, FG_TITLE, RADIUS,
)
from bin_parser import quick_scan_bin, process_multiple_files
from log_parser import process_log_files
from centralizator import update_centralizator, update_centralizator_batch
from excel_report import reset_ddp_perf, get_ddp_perf_seconds
from contoare_db import _load_contoare_db, _save_contoare_db, _delete_contor_from_centralizator
from gui_widgets import (
    _rr, _ctk_btn, _make_button, _ctk_frame, _ctk_label,
    _ctk_entry, _add_logo_header, _open_contor_dialog,
)

from harta_server import HartaServer

def _format_perf_summary(phases, total_seconds):
    """
    Genereaza un raport scurt: unde s-a dus timpul in timpul procesarii.
    phases: lista de tuple (eticheta, secunde)
    """
    if not phases or total_seconds <= 0:
        return "     (fara date de performanta)"
    max_label = max(len(lbl) for lbl, _ in phases)
    linii = []
    for lbl, secs in phases:
        pct = (secs / total_seconds) * 100 if total_seconds else 0
        linii.append(f"     {lbl.ljust(max_label)}   {secs:7.1f}s   ({pct:5.1f}%)")
    linii.append(f"     {'TOTAL'.ljust(max_label)}   {total_seconds:7.1f}s   (100.0%)")
    return "\n".join(linii)


class PeekApp(ctk.CTk if CTK_AVAILABLE else tk.Tk):

    def __init__(self):
        # ── DPI awareness ─────────────────────────────────────────────────────
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        if CTK_AVAILABLE:
            # Detectăm light/dark din registry Windows
            try:
                import winreg
                reg = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                val, _ = winreg.QueryValueEx(reg, "AppsUseLightTheme")
                winreg.CloseKey(reg)
                detected_theme = "light" if val == 1 else "dark"
            except Exception:
                detected_theme = "light"

            ctk.set_appearance_mode(detected_theme)

            # Calea la blue.json — în executabil PyInstaller customtkinter
            # e copiat direct în _MEIPASS/customtkinter/ (nu în site-packages)
            try:
                if getattr(sys, 'frozen', False):
                    ctk_dir = os.path.join(sys._MEIPASS, 'customtkinter')
                else:
                    import customtkinter as _ctk_mod
                    ctk_dir = os.path.dirname(_ctk_mod.__file__)

                theme_file = os.path.join(ctk_dir, 'assets', 'themes', 'blue.json')
                if os.path.isfile(theme_file):
                    ctk.set_default_color_theme(theme_file)
                else:
                    # fallback — listăm ce avem pentru debug
                    themes_dir = os.path.join(ctk_dir, 'assets', 'themes')
                    if os.path.isdir(themes_dir):
                        files = os.listdir(themes_dir)
                        if files:
                            ctk.set_default_color_theme(
                                os.path.join(themes_dir, files[0]))
                        else:
                            ctk.set_default_color_theme('blue')
                    else:
                        ctk.set_default_color_theme('blue')
            except Exception:
                ctk.set_default_color_theme('blue')

        super().__init__()
        self.title("PEEK Traffic Analyzer")
        self.resizable(True, True)
        self.minsize(520, 600)

        W, H = 660, 760
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        self.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        # Colturi rotunjite Windows 11
        self.after(100, self._try_round)

        if getattr(sys, 'frozen', False):
            self.base_path = sys._MEIPASS
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))

        self.selected_files    = []
        self.selected_log_files = []
        self.stop_event        = threading.Event()
        self.processing_thread = None

        self._setup_ui()

        # ── Server hartă (Flask, thread daemon) ──────────────────────────────
        self._harta_server = HartaServer()
        self._harta_server.start()

        # ── Sincronizare contoare.db la pornire ───────────────────────────────
        # Dacă contoare.db e gol și Excel-ul centralizator are date,
        # importăm automat în fundal (o singură dată, non-blocking)
        threading.Thread(target=self._sync_contoare_db_on_start,
                         daemon=True).start()

    def _try_round(self):
        try:
            import ctypes
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                self.winfo_id(), 33, ctypes.byref(ctypes.c_int(2)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

    def _sync_contoare_db_on_start(self):
        """
        Sincronizare automată la pornire: dacă contoare.db e gol
        dar Excel-ul centralizator are date, importăm toate contoarele.
        Rulează în thread daemon — nu blochează GUI.
        """
        try:
            from database import get_contoare_db
            from contoare_db import _load_contoare_db

            cdb = get_contoare_db()
            # Verificăm dacă DB-ul e gol
            toate = cdb.get_all()
            if toate:
                return   # DB deja populat, nimic de făcut

            # DB gol → importăm din Excel
            db_excel = _load_contoare_db()
            if not db_excel:
                return   # nici Excel-ul nu are date

            n = 0
            for ct_id, data in db_excel.items():
                cdb.upsert(ct_id, data)
                n += 1

            if n > 0:
                self.after(0, lambda: self._log(
                    f"  🗄️  contoare.db sincronizat automat: {n} contoare importate din Excel."))
        except Exception as _e:
            print(f"[WARN] Sincronizare contoare.db eșuată: {_e}")

    def _setup_ui(self):
        # ── Iconita fereastra (bara de titlu Windows) ─────────────────────────
        ico_path  = os.path.join(self.base_path, "assets", "icon.ico")
        png_path  = os.path.join(self.base_path, "assets", "icon.png")
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass
        elif PIL_AVAILABLE and os.path.exists(png_path):
            try:
                _ico_img = Image.open(png_path).resize((32, 32), Image.Resampling.LANCZOS)
                self._tk_icon = ImageTk.PhotoImage(_ico_img)
                self.iconphoto(True, self._tk_icon)
            except Exception:
                pass

        # ══════════════════════════════════════════════════════════════════════
        # Layout principal: rând superior cu 2 coloane egale (PEEK | VEK)
        # ══════════════════════════════════════════════════════════════════════
        if CTK_AVAILABLE:
            top_row = ctk.CTkFrame(self, fg_color="transparent")
        else:
            top_row = tk.Frame(self)
        top_row.pack(fill="x", padx=16, pady=(12, 4))
        top_row.columnconfigure(0, weight=1)
        top_row.columnconfigure(1, weight=1)

        # ── COLOANA STANGA: PEEK ─────────────────────────────────────────────
        if CTK_AVAILABLE:
            card_peek = ctk.CTkFrame(top_row, fg_color="#FFFFFF", corner_radius=10)
        else:
            card_peek = tk.Frame(top_row, bg="#FFFFFF", relief="solid", bd=1)
        card_peek.grid(row=0, column=0, sticky="nsew", padx=(0, 6), ipady=8)

        # Logo PEEK
        peek_logo_path = None
        for fname in ("logo.png", "peek_logo.png", "icon.png"):
            p = os.path.join(self.base_path, "assets", fname)
            if os.path.exists(p):
                peek_logo_path = p
                break

        if PIL_AVAILABLE and peek_logo_path:
            try:
                img = Image.open(peek_logo_path)
                img.thumbnail((220, 70), Image.Resampling.LANCZOS)
                if CTK_AVAILABLE:
                    self._logo_peek = ctk.CTkImage(
                        light_image=img, dark_image=img,
                        size=(img.width, img.height))
                    ctk.CTkLabel(card_peek, image=self._logo_peek, text="",
                                 fg_color="transparent").pack(pady=(10, 4))
                else:
                    self._logo_peek = ImageTk.PhotoImage(img)
                    tk.Label(card_peek, image=self._logo_peek,
                             bg="#FFFFFF").pack(pady=(10, 4))
            except Exception:
                self._logo_fallback(card_peek)
        else:
            self._logo_fallback(card_peek)

        _ctk_label(card_peek, "Fisiere .bin",
                   font=("Segoe UI", 10, "bold"),
                   text_color="#555555").pack(pady=(4, 2))

        self.btn_files = _ctk_btn(card_peek, "📂  Alege fisiere .bin",
                                  self._choose_files, "primary", width=200)
        self.btn_files.pack(pady=4)

        self.btn_folder_bin = _ctk_btn(card_peek, "📁  Folder recursiv .bin",
                                       self._choose_folder_bin, "navy", width=200)
        self.btn_folder_bin.pack(pady=(0, 4))

        self.lbl_files = _ctk_label(card_peek, "Niciun fișier .bin selectat.",
                                    font=FONT_SMALL, text_color="#888888")
        self.lbl_files.pack(pady=(2, 8))

        # ── COLOANA DREAPTA: VEK ─────────────────────────────────────────────
        if CTK_AVAILABLE:
            card_vek = ctk.CTkFrame(top_row, fg_color="#FFFFFF", corner_radius=10)
        else:
            card_vek = tk.Frame(top_row, bg="#FFFFFF", relief="solid", bd=1)
        card_vek.grid(row=0, column=1, sticky="nsew", padx=(6, 0), ipady=8)

        # Logo VEK
        vek_logo_path = None
        for fname in ("icon_vek.png", "vek_logo.png", "vek_icon.png"):
            p = os.path.join(self.base_path, "assets", fname)
            if os.path.exists(p):
                vek_logo_path = p
                break

        if PIL_AVAILABLE and vek_logo_path:
            try:
                img_v = Image.open(vek_logo_path)
                img_v.thumbnail((220, 70), Image.Resampling.LANCZOS)
                if CTK_AVAILABLE:
                    self._logo_vek = ctk.CTkImage(
                        light_image=img_v, dark_image=img_v,
                        size=(img_v.width, img_v.height))
                    ctk.CTkLabel(card_vek, image=self._logo_vek, text="",
                                 fg_color="transparent").pack(pady=(10, 4))
                else:
                    self._logo_vek = ImageTk.PhotoImage(img_v)
                    tk.Label(card_vek, image=self._logo_vek,
                             bg="#FFFFFF").pack(pady=(10, 4))
            except Exception:
                _ctk_label(card_vek, "VEK Traffic Analyzer",
                           font=("Segoe UI", 13, "bold"),
                           text_color="#1A5276").pack(pady=(10, 4))
        else:
            _ctk_label(card_vek, "VEK Traffic Analyzer",
                       font=("Segoe UI", 13, "bold"),
                       text_color="#1A5276").pack(pady=(10, 4))

        _ctk_label(card_vek, "Fisiere .log",
                   font=("Segoe UI", 10, "bold"),
                   text_color="#555555").pack(pady=(4, 2))

        self.btn_log_files = _ctk_btn(card_vek, "📂  Alege fisiere .log",
                                      self._choose_log_files, "primary", width=200)
        self.btn_log_files.pack(pady=4)

        self.btn_folder_log = _ctk_btn(card_vek, "📁  Folder recursiv .log",
                                       self._choose_folder_log, "navy", width=200)
        self.btn_folder_log.pack(pady=(0, 4))

        self.lbl_log_files = _ctk_label(card_vek, "Niciun fișier .log selectat.",
                                        font=FONT_SMALL, text_color="#888888")
        self.lbl_log_files.pack(pady=(2, 8))

        # ══════════════════════════════════════════════════════════════════════
        # Buton unic de procesare (proceseaza .bin si/sau .log selectate)
        # ══════════════════════════════════════════════════════════════════════
        if CTK_AVAILABLE:
            row_btns = ctk.CTkFrame(self, fg_color="transparent")
        else:
            row_btns = tk.Frame(self)
        row_btns.pack(pady=(6, 2))

        self.btn_process = _ctk_btn(row_btns, "▶  Proceseaza fisierele",
                                    self._run_all_processing, "success", width=230)
        self.btn_process.pack(side="left", padx=8)

        self.btn_cancel = _ctk_btn(row_btns, "⏹  Anuleaza",
                                   self._cancel_processing, "danger",
                                   width=150, state="disabled")
        self.btn_cancel.pack(side="left", padx=8)

        self.btn_force_unlock = _ctk_btn(row_btns, "🔓  Eliberează lock",
                                         self._force_release_lock, "warning",
                                         width=170, state="normal")
        self.btn_force_unlock.pack(side="left", padx=8)

        # ── Butoane acțiuni secundare ─────────────────────────────────────────
        if CTK_AVAILABLE:
            row_actions = ctk.CTkFrame(self, fg_color="transparent")
        else:
            row_actions = tk.Frame(self)
        row_actions.pack(pady=(4, 6))

        self.btn_contoare = _ctk_btn(row_actions, "🗄️  Gestionare Contoare",
                                     self._open_contoare_manager, "navy", width=220)
        self.btn_contoare.pack(side="left", padx=6)

        self.btn_mzl_manual = _ctk_btn(row_actions, "✏️  Procesare manuală MZL",
                                        self._open_mzl_manual_dialog, "info", width=220)
        self.btn_mzl_manual.pack(side="left", padx=6)

        # ── Al doilea rând de acțiuni ─────────────────────────────────────────
        if CTK_AVAILABLE:
            row_actions2 = ctk.CTkFrame(self, fg_color="transparent")
        else:
            row_actions2 = tk.Frame(self)
        row_actions2.pack(pady=(0, 4))

        self.btn_sorteaza = _ctk_btn(row_actions2, "📦  Sortează .bin după post",
                                     self._sorteaza_bin_dialog, "warning", width=230)
        self.btn_sorteaza.pack(side="left", padx=6)

        self.btn_harta = _ctk_btn(row_actions2, "🗺️  Hartă Contoare",
                                  self._open_harta, "navy", width=190)
        self.btn_harta.pack(side="left", padx=6)

        # ── Progress bar ──────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            self.progress = ctk.CTkProgressBar(self, width=560, height=14,
                                               corner_radius=7,
                                               progress_color="#1A7A3C")
            self.progress.set(0)
            self.progress.pack(pady=(8, 2))
        else:
            self.progress = ttk.Progressbar(self, orient="horizontal",
                                            length=540, mode="determinate")
            self.progress.pack(pady=(8, 2))

        self.lbl_status = _ctk_label(self, "Gata de procesare.",
                                     font=FONT_SMALL, text_color="#555555")
        self.lbl_status.pack(pady=(0, 4))

        # ── Jurnal procesare ──────────────────────────────────────────────────
        log_card = _ctk_frame(self)
        log_card.pack(fill="both", expand=True, padx=20, pady=(4, 16))

        _ctk_label(log_card, "  Jurnal procesare",
                   font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8, 2))

        if CTK_AVAILABLE:
            self.log_text = ctk.CTkTextbox(
                log_card, font=FONT_MONO,
                corner_radius=6, wrap="word",
                state="disabled")
            self.log_text.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        else:
            txt_frame = tk.Frame(log_card)
            txt_frame.pack(fill="both", expand=True, padx=8, pady=(2, 8))
            self.log_text = tk.Text(txt_frame, font=FONT_MONO,
                                    state="disabled", relief="flat", wrap="word",
                                    bg="#F8F9FA", fg="#212529")
            sb = ttk.Scrollbar(txt_frame, orient="vertical",
                               command=self.log_text.yview)
            self.log_text.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            self.log_text.pack(fill="both", expand=True)

        self._log("Aplicatie pornita. Selecteaza fisierele .bin si/sau .log pentru procesare.")

    def _logo_fallback(self, parent=None):
        p = parent or self
        _ctk_label(p, "PEEK Traffic Analyzer",
                   font=("Segoe UI", 20, "bold"),
                   text_color="#C0392B").pack(pady=(20, 6))

    def _open_harta(self):
        """Deschide harta contoare în browser-ul default."""
        self._harta_server.open_in_browser()

    # ══════════════════════════════════════════════════════════════════════════
    # Sortare .bin după post
    # ══════════════════════════════════════════════════════════════════════════
    def _sorteaza_bin_dialog(self):
        """Deschide fereastra de configurare și progres pentru sortarea .bin."""
        win = ctk.CTkToplevel(self) if CTK_AVAILABLE else tk.Toplevel(self)
        win.title("Sortare fișiere .bin după post")
        win.resizable(True, True)
        win.grab_set()

        W, H = 680, 520
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        win.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        # ── Cai configurabile ─────────────────────────────────────────────────
        DEFAULT_SURSA = r"L:\BIDMRCT\datePEEK\Descarcari"
        DEFAULT_DEST  = r"L:\BIDMRCT\datePEEK"

        var_sursa = tk.StringVar(value=DEFAULT_SURSA)
        var_dest  = tk.StringVar(value=DEFAULT_DEST)

        if CTK_AVAILABLE:
            hdr = ctk.CTkFrame(win, fg_color="#1A5276", corner_radius=0)
        else:
            hdr = tk.Frame(win, bg="#1A5276")
        hdr.pack(fill="x")
        _ctk_label(hdr, "  📦  Sortare fișiere .bin după post",
                   font=("Segoe UI", 13, "bold"),
                   text_color="#FFFFFF").pack(anchor="w", padx=12, pady=10)

        # ── Formular cai ──────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            form = ctk.CTkFrame(win, fg_color="transparent")
        else:
            form = tk.Frame(win)
        form.pack(fill="x", padx=16, pady=10)

        def _row_cale(parent, label, var, row):
            if CTK_AVAILABLE:
                ctk.CTkLabel(parent, text=label, font=FONT_SMALL,
                             anchor="e", width=120).grid(
                    row=row, column=0, padx=(0, 8), pady=6, sticky="e")
                e = ctk.CTkEntry(parent, textvariable=var, width=380,
                                 height=32, corner_radius=6)
                e.grid(row=row, column=1, sticky="ew", pady=6)
                btn = ctk.CTkButton(parent, text="📁", width=36, height=32,
                                    corner_radius=6,
                                    command=lambda v=var: _browse(v))
                btn.grid(row=row, column=2, padx=(6, 0), pady=6)
            else:
                tk.Label(parent, text=label, font=FONT_SMALL,
                         anchor="e", width=16).grid(
                    row=row, column=0, padx=(0, 8), pady=6, sticky="e")
                e = ttk.Entry(parent, textvariable=var, width=48)
                e.grid(row=row, column=1, sticky="ew", pady=6)
                tk.Button(parent, text="📁",
                          command=lambda v=var: _browse(v)).grid(
                    row=row, column=2, padx=(6, 0), pady=6)
            parent.columnconfigure(1, weight=1)

        def _browse(var):
            folder = filedialog.askdirectory(parent=win)
            if folder:
                var.set(folder)

        _row_cale(form, "Folder sursă:", var_sursa, 0)
        _row_cale(form, "Folder posturi:", var_dest,  1)

        _ctk_label(form,
                   "  ℹ️  Fișierele .bin din sursă vor fi mutate în subfolderele\n"
                   "      de forma  XXXX_Localitate  din folderul posturi.",
                   font=FONT_SMALL, text_color="#555555").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 4))

        # ── Log fereastră ─────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            log_frame = ctk.CTkFrame(win, fg_color="#F8F9FA", corner_radius=8)
        else:
            log_frame = tk.Frame(win, relief="sunken", bd=1)
        log_frame.pack(fill="both", expand=True, padx=16, pady=(4, 8))

        if CTK_AVAILABLE:
            log_box = ctk.CTkTextbox(log_frame, font=FONT_MONO,
                                     corner_radius=6, wrap="word",
                                     state="disabled")
            log_box.pack(fill="both", expand=True, padx=6, pady=6)
        else:
            txt_fr = tk.Frame(log_frame)
            txt_fr.pack(fill="both", expand=True, padx=6, pady=6)
            log_box = tk.Text(txt_fr, font=FONT_MONO, state="disabled",
                              relief="flat", wrap="word",
                              bg="#F8F9FA", fg="#212529")
            sb2 = ttk.Scrollbar(txt_fr, orient="vertical",
                                 command=log_box.yview)
            log_box.configure(yscrollcommand=sb2.set)
            sb2.pack(side="right", fill="y")
            log_box.pack(fill="both", expand=True)

        def _log_win(msg):
            if CTK_AVAILABLE:
                log_box.configure(state="normal")
                log_box.insert("end", msg + "\n")
                log_box.see("end")
                log_box.configure(state="disabled")
            else:
                log_box.config(state="normal")
                log_box.insert("end", msg + "\n")
                log_box.see("end")
                log_box.config(state="disabled")

        # ── Butoane ───────────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            btn_row = ctk.CTkFrame(win, fg_color="transparent")
        else:
            btn_row = tk.Frame(win)
        btn_row.pack(pady=(0, 12))

        btn_start = _ctk_btn(btn_row, "▶  Pornește sortarea",
                             lambda: _start(), "success", width=200)
        btn_start.pack(side="left", padx=8)

        _ctk_btn(btn_row, "✖  Închide", win.destroy,
                 "secondary", width=130).pack(side="left", padx=8)

        # ── Logica de sortare (rulează în thread) ─────────────────────────────
        def _start():
            sursa = var_sursa.get().strip()
            dest  = var_dest.get().strip()

            if not os.path.isdir(sursa):
                messagebox.showerror("Eroare",
                    f"Folderul sursă nu există:\n{sursa}", parent=win)
                return
            if not os.path.isdir(dest):
                messagebox.showerror("Eroare",
                    f"Folderul posturi nu există:\n{dest}", parent=win)
                return

            btn_start.configure(state="disabled")
            # Șterge log-ul anterior
            if CTK_AVAILABLE:
                log_box.configure(state="normal")
                log_box.delete("0.0", "end")
                log_box.configure(state="disabled")
            else:
                log_box.config(state="normal")
                log_box.delete("1.0", "end")
                log_box.config(state="disabled")

            threading.Thread(
                target=_run_sortare,
                args=(sursa, dest),
                daemon=True
            ).start()

        def _run_sortare(sursa, dest):
            import shutil

            def _extrage_site_id(filepath):
                try:
                    with open(filepath, "rb") as f:
                        raw = f.read(150)
                    header_text = "".join(
                        [chr(b) if 32 <= b <= 126 else " " for b in raw])
                    matches = re.findall(r'000+(\d{4})', header_text)
                    if matches:
                        return matches[-1]
                    base_name = os.path.basename(filepath)
                    m = re.search(r'(\d{4})', base_name)
                    return m.group(1) if m else None
                except Exception as e:
                    return None

            # Construim map cod -> folder destinatie
            folder_map = {}
            for entry in os.scandir(dest):
                if entry.is_dir():
                    m = re.match(r'^(\d{4})_', entry.name)
                    if m:
                        folder_map[m.group(1)] = entry.path

            win.after(0, lambda: _log_win(
                f"{'─'*55}\n"
                f"  Sursă:    {sursa}\n"
                f"  Posturi:  {dest}\n"
                f"  Foldere posturi găsite: {len(folder_map)}\n"
                f"{'─'*55}"))

            fisiere = [f for f in os.listdir(sursa) if f.lower().endswith(".bin")]
            if not fisiere:
                win.after(0, lambda: _log_win("  ⚠️  Nu există fișiere .bin în folderul sursă."))
                win.after(0, lambda: btn_start.configure(state="normal"))
                return

            win.after(0, lambda: _log_win(f"  Fișiere .bin găsite: {len(fisiere)}\n"))

            mutate = 0; suprascrise = 0; negasite = 0; erori = 0

            for fisier in sorted(fisiere):
                cale_sursa = os.path.join(sursa, fisier)
                site_id = _extrage_site_id(cale_sursa)

                if site_id is None:
                    msg = f"  ⚠️  {fisier}  →  cod neidentificat, omis"
                    win.after(0, lambda m=msg: _log_win(m))
                    negasite += 1
                    continue

                if site_id not in folder_map:
                    msg = f"  ❓  {fisier}  →  cod {site_id} fără folder corespondent"
                    win.after(0, lambda m=msg: _log_win(m))
                    negasite += 1
                    continue

                folder_dest = folder_map[site_id]
                cale_dest   = os.path.join(folder_dest, fisier)
                exista_deja = os.path.exists(cale_dest)

                try:
                    shutil.move(cale_sursa, cale_dest)
                    if exista_deja:
                        msg = f"  🔄  {fisier}  →  {os.path.basename(folder_dest)}  (suprascriere)"
                        suprascrise += 1
                    else:
                        msg = f"  ✅  {fisier}  →  {os.path.basename(folder_dest)}"
                        mutate += 1
                    win.after(0, lambda m=msg: _log_win(m))
                    self.after(0, lambda m=msg: self._log(m))
                except Exception as e:
                    msg = f"  ❌  {fisier}  →  eroare: {e}"
                    win.after(0, lambda m=msg: _log_win(m))
                    erori += 1

            sumar = (
                f"\n{'─'*55}\n"
                f"  ✅  Mutate cu succes:   {mutate}\n"
                f"  🔄  Suprascrise:        {suprascrise}\n"
                f"  ❓  Fără folder:        {negasite}\n"
                f"  ❌  Erori:              {erori}\n"
                f"{'─'*55}"
            )
            win.after(0, lambda: _log_win(sumar))
            self.after(0, lambda: self._log(
                f"  📦  Sortare .bin finalizată — "
                f"mutate: {mutate}, suprascrise: {suprascrise}, "
                f"fără folder: {negasite}, erori: {erori}"))
            win.after(0, lambda: btn_start.configure(state="normal"))

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _update_progress(self, percent, text):
        def _do():
            if CTK_AVAILABLE:
                self.progress.set(percent / 100)
            else:
                self.progress["value"] = percent
            self.lbl_status.configure(text=text)
        self.after(0, _do)

    def _log(self, msg):
        def _do():
            if CTK_AVAILABLE:
                self.log_text.configure(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            else:
                self.log_text.config(state="normal")
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
                self.log_text.config(state="disabled")
        self.after(0, _do)

    # ── Selectare folder recursiv .bin ────────────────────────────────────────
    def _choose_folder_bin(self):
        folder = filedialog.askdirectory(
            title="Selectează folderul rădăcină — caută recursiv fișiere .bin",
        )
        if not folder:
            return
        paths = []
        for root_dir, dirs, files in os.walk(folder):
            dirs.sort()
            for fname in sorted(files):
                if fname.lower().endswith(".bin"):
                    paths.append(os.path.join(root_dir, fname))
        if not paths:
            messagebox.showwarning(
                "Niciun fișier găsit",
                f"Nu s-au găsit fișiere .bin în:\n{folder}",
            )
            return
        self.selected_files = paths
        n = len(paths)
        def _sid(p):
            base = os.path.splitext(os.path.basename(p))[0]
            m = re.findall(r'000+(\d{4})', base)
            if m: return m[-1]
            m2 = re.findall(r'(\d{4})', base)
            return m2[-1] if m2 else "????"
        ct = sorted(set(_sid(p) for p in paths))
        nc = len(ct)
        self.lbl_files.configure(
            text=f"{n} fișier{'e' if n!=1 else ''} .bin (recursiv) │ {nc} contoar{'e' if nc!=1 else ''}.")
        sep = "─" * 52
        self._log(f"\n{sep}")
        self._log(f"  📁 Folder recursiv .bin: {folder}")
        self._log(f"  📂 Fișiere: {n}  │  Contoar{'e' if nc!=1 else ''}: {nc}")
        if nc <= 20: self._log(f"  📋 {', '.join(ct)}")
        else:        self._log(f"  📋 {', '.join(ct[:20])} ... (+{nc-20})")
        self._log(sep)

    # ── Selectare folder recursiv .log ────────────────────────────────────────
    def _choose_folder_log(self):
        folder = filedialog.askdirectory(
            title="Selectează folderul rădăcină — caută recursiv fișiere .log",
        )
        if not folder:
            return
        paths = []
        for root_dir, dirs, files in os.walk(folder):
            dirs.sort()
            for fname in sorted(files):
                if fname.lower().endswith(".log"):
                    paths.append(os.path.join(root_dir, fname))
        if not paths:
            messagebox.showwarning(
                "Niciun fișier găsit",
                f"Nu s-au găsit fișiere .log în:\n{folder}",
            )
            return
        self.selected_log_files = paths
        n = len(paths)
        def _sid(p):
            base = os.path.splitext(os.path.basename(p))[0]
            return base.split("_")[0]
        ct = sorted(set(_sid(p) for p in paths))
        nc = len(ct)
        self.lbl_log_files.configure(
            text=f"{n} fișier{'e' if n!=1 else ''} .log (recursiv) │ {nc} contoar{'e' if nc!=1 else ''}.")
        sep = "─" * 52
        self._log(f"\n{sep}")
        self._log(f"  📁 Folder recursiv .log: {folder}")
        self._log(f"  📂 Fișiere: {n}  │  Contoar{'e' if nc!=1 else ''}: {nc}")
        if nc <= 20: self._log(f"  📋 {', '.join(ct)}")
        else:        self._log(f"  📋 {', '.join(ct[:20])} ... (+{nc-20})")
        self._log(sep)

    # ── Callbacks ─────────────────────────────────────────────────────────────
    def _choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Selectează fișiere .bin PEEK",
            filetypes=[("Fișiere .bin", "*.bin"), ("Toate fișierele", "*.*")],
        )
        if not paths:
            return
        self.selected_files = list(paths)
        n = len(paths)
        def _sid(p):
            base = os.path.splitext(os.path.basename(p))[0]
            m = re.findall(r'000+(\d{4})', base)
            if m: return m[-1]
            m2 = re.findall(r'(\d{4})', base)
            return m2[-1] if m2 else "????"
        ct = sorted(set(_sid(p) for p in paths))
        nc = len(ct)
        self.lbl_files.configure(
            text=f"{n} fișier{'e' if n!=1 else ''} .bin selectat{'e' if n!=1 else ''} "
                 f"de la {nc} contoar{'e' if nc!=1 else ''}.")
        sep = "─" * 52
        self._log(f"\n{sep}")
        self._log(f"  📂 Fișier{'e' if n!=1 else ''} .bin: {n}  │  Contoar{'e' if nc!=1 else ''}: {nc}")
        if nc <= 20: self._log(f"  📋 {', '.join(ct)}")
        else:        self._log(f"  📋 {', '.join(ct[:20])} ... (+{nc-20})")
        self._log(sep)

    # ── Procesare unificata (.bin + .log) ─────────────────────────────────────
    # ══════════════════════════════════════════════════════════════════════════
    # Lock file — procesare exclusivă pe rețea (bin + MZL manual)
    # ══════════════════════════════════════════════════════════════════════════
    # Timeout de siguranță: 30 minute. Dacă aplicația cade în mijlocul
    # unei operații, lock-ul e ignorat automat după 30 minute.
    # Lock-ul normal se eliberează imediat în blocul finally al fiecărei operații.
    _LOCK_TIMEOUT = 1800  # 30 minute

    def _get_lock_path(self):
        """Calea fișierului .procesare.lock pe drive-ul comun."""
        try:
            return os.path.join(CENTRAL_FILE_FOLDER, ".procesare.lock")
        except Exception:
            return None

    def _acquire_lock(self, tip_operatie: str = "Operație"):
        """
        Încearcă să obțină lock-ul.
        tip_operatie: text afișat celorlalți utilizatori (ex: 'Procesare .bin (5 fișiere)')
        Returnează True dacă a reușit, 'OCUPAT:info' dacă e blocat.
        """
        import getpass, socket
        lock_path = self._get_lock_path()
        if lock_path is None:
            return True
        try:
            if os.path.exists(lock_path):
                age = time.time() - os.path.getmtime(lock_path)
                if age > self._LOCK_TIMEOUT:
                    # Lock vechi — expirat, îl ștergem și continuăm
                    try:
                        os.remove(lock_path)
                        self._log("⚠  Lock vechi detectat și eliberat automat (timeout 30 min).")
                    except Exception:
                        return True
                else:
                    try:
                        info = open(lock_path, encoding="utf-8").read().strip()
                    except Exception:
                        info = "alt utilizator (detalii indisponibile)"
                    return f"OCUPAT:{info}"
            # Scriem lock-ul nostru
            user = getpass.getuser()
            host = socket.gethostname()
            ts   = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            info = f"{user}@{host} | {ts} | {tip_operatie}"
            with open(lock_path, "w", encoding="utf-8") as f:
                f.write(info)
            return True
        except Exception:
            return True  # dacă nu putem gestiona lock-ul, continuăm oricum

    def _release_lock(self):
        """Eliberează lock-ul. Apelat în finally — rulează întotdeauna."""
        lock_path = self._get_lock_path()
        if not lock_path:
            return
        try:
            if os.path.exists(lock_path):
                os.remove(lock_path)
        except Exception:
            pass

    def _check_lock_and_warn(self, tip_operatie: str = "Operație") -> bool:
        """
        Verifică lock-ul și afișează mesaj dacă e ocupat.
        Returnează True dacă putem continua, False dacă e blocat.
        """
        result = self._acquire_lock(tip_operatie)
        if isinstance(result, str) and result.startswith("OCUPAT:"):
            info = result[7:]
            messagebox.showwarning(
                "Operație în curs",
                f"Un alt utilizator efectuează o operație în acest moment:\n\n"
                f"  {info}\n\n"
                f"Așteptați finalizarea și încercați din nou.\n"
                f"Dacă operația s-a terminat și mesajul persistă,\n"
                f"apăsați butonul  🔓 Eliberează lock.",
                parent=self,
            )
            return False
        return True

    def _run_all_processing(self):
        has_bin = bool(self.selected_files)
        has_log = bool(self.selected_log_files)
        if not has_bin and not has_log:
            messagebox.showwarning("Atentie", "Selecteaza cel putin un fisier .bin sau .log!")
            return

        # ── Verificare lock rețea ─────────────────────────────────────────────
        n_bin = len(self.selected_files)
        n_log = len(self.selected_log_files)
        tip   = f"Procesare fișiere ({n_bin} .bin" + (f", {n_log} .log" if n_log else "") + ")"
        if not self._check_lock_and_warn(tip):
            return

        self.stop_event.clear()
        self.btn_process.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        msgs = []
        if has_bin: msgs.append(f"{len(self.selected_files)} .bin")
        if has_log: msgs.append(f"{len(self.selected_log_files)} .log")
        self._log(f"\n── Pornesc procesarea: {', '.join(msgs)} ─────────────────────")
        self.processing_thread = threading.Thread(
            target=self._all_processing_background, daemon=True)
        self.processing_thread.start()

    def _run_processing(self):
        """Pastrat pentru compatibilitate interna — redirectat la _run_all_processing."""
        self._run_all_processing()
        if not self.selected_files:
            messagebox.showwarning("Atenție", "Selectează cel puțin un fișier .bin!")
            return
        self.stop_event.clear()
        self.btn_process.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self._log("\n── Pornesc procesarea ──────────────────────────────────")
        self.processing_thread = threading.Thread(
            target=self._process_in_background, daemon=True)
        self.processing_thread.start()

    def _process_in_background(self):
        total = len(self.selected_files)
        sep   = "─" * 52

        def cancelled():
            """Returneaza True daca utilizatorul a apasat Anuleaza."""
            return self.stop_event.is_set()

        def abort(msg="⛔ Procesare anulată de utilizator."):
            self._log(f"\n{msg}")
            self._update_progress(0, "⛔ Anulat.")

        try:
            self._update_progress(2, f"Se scanează {total} fișiere .bin...")
            self._log(f"\n{sep}\n  🔍 Scanare: {total} fișiere\n{sep}")

            scan_results = []; fisiere_goale = 0
            for idx, p in enumerate(self.selected_files, 1):
                if cancelled(): abort(); return          # ← verificare anulare
                sid, n_brut, n_lanes = quick_scan_bin(p)
                scan_results.append((p, sid, n_brut, n_lanes))
                fname = os.path.basename(p)
                pct   = int(2 + (idx / total) * 8)
                if n_brut == 0:
                    self._log(f"  ⚠  [{sid}]  {fname}  →  fișier gol")
                    fisiere_goale += 1
                else:
                    self._log(f"  ✓  [{sid}]  {fname}  →  ~{n_brut} înreg.  │  {n_lanes} benzi")
                self._update_progress(pct, f"Scanare {idx}/{total}...")

            if cancelled(): abort(); return              # ← după scanare

            fisiere_valide = total - fisiere_goale
            goale_str = f"  ({fisiere_goale} goale)" if fisiere_goale else "  (toate valide)"
            self._log(f"\n  📋 Scanare: {fisiere_valide}/{total} valide{goale_str}\n")
            self._update_progress(10, "Se procesează datele...")
            self._log("  🔄 Generare rapoarte Excel...")

            rezultate = process_multiple_files(self.selected_files, stop_event=self.stop_event)
            if rezultate is None and cancelled():
                abort(); return
            if not rezultate:
                self._log("\n✗ Nicio dată validă găsită.")
                self._update_progress(0, "✗  Nicio dată validă.")
                return

            if cancelled(): abort(); return              # ← după procesare Excel

            n_contoare = len(rezultate)
            total_ore  = sum(r['randuri'] for r in rezultate)
            total_b1   = sum(r['b1'] for r in rezultate)
            total_b2   = sum(r['b2'] for r in rezultate)

            self._update_progress(75, f"Actualizez centralizatorul ({n_contoare} contoare)...")
            if cancelled(): abort(); return          # ← înainte de centralizator
            _, _erori_c = update_centralizator_batch(
                rezultate, CENTRAL_FILE_FOLDER, log_callback=self._log)
            self._update_progress(95, "Centralizator actualizat.")

            if cancelled(): abort(); return              # ← după centralizator

            lanes_info = ", ".join(f"{r['id']}:{r['n_lanes']}b" for r in rezultate)
            self._log(f"\n{sep}\n  📦 TOTAL GENERAL:")
            self._log(f"     Contoare procesate:  {n_contoare}")
            self._log(f"     Fișiere:             {fisiere_valide}/{total}{goale_str}")
            self._log(f"     Înregistrări orare:  {total_ore:,}")
            self._log(f"     Total trafic:        {int(total_b1+total_b2):,}")
            self._log(f"     Benzi per contoar:    {lanes_info}")
            self._log(f"{sep}\n")

            self._update_progress(100, f"✔  Finalizat — {n_contoare} contoar{'e' if n_contoare!=1 else ''} procesat{'e' if n_contoare!=1 else ''}.")
            rapoarte_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)
            messagebox.showinfo(
                "Succes",
                f"Procesare finalizată!\n\n"
                f"Contoar{'e' if n_contoare!=1 else ''} procesat{'e' if n_contoare!=1 else ''}:   {n_contoare}\n"
                f"Fișiere:              {fisiere_valide}/{total}{goale_str}\n"
                f"Înregistrări orare:   {total_ore:,}\n\n"
                f"Rapoartele au fost salvate în:\n{rapoarte_dir}",
            )
        except Exception as e:
            import traceback
            self._log(f"\n✗ EROARE: {e}")
            traceback.print_exc()
            self._update_progress(0, "✗  Eroare la procesare.")
        finally:
            self._release_lock()
            self.btn_process.configure(state="normal")
            self.btn_cancel.configure(state="disabled")
            self.selected_files = []
            self.lbl_files.configure(text="Niciun fișier .bin selectat.")
        """
        Thread background unificat: proceseaza .bin si/sau .log in ordine.
        Reutilizeaza logica din _process_in_background si _log_process_in_background.
        """
        sep = "─" * 52

        def cancelled():
            return self.stop_event.is_set()

        def abort(msg="⛔ Procesare anulata de utilizator."):
            self._log(f"\n{msg}")
            self._update_progress(0, "⛔ Anulat.")
            self.after(0, lambda: self.btn_process.configure(state="normal"))
            self.after(0, lambda: self.btn_cancel.configure(state="disabled"))

        try:
            has_bin = bool(self.selected_files)
            has_log = bool(self.selected_log_files)

            # Calculam procentele: daca avem ambele, impartim 0-70 pt bin, 70-100 pt log
            # Daca avem doar unul, folosim 0-100
            bin_pct_end = 95 if not has_log else 60
            log_pct_start = bin_pct_end if has_bin else 5

            rezultate_bin = []
            rezultate_log = []

            # ═══════════════════════════════════════════════════════════════
            # BLOC .BIN
            # ═══════════════════════════════════════════════════════════════
            if has_bin:
                total = len(self.selected_files)
                self._update_progress(2, f"Se scaneaza {total} fisiere .bin...")
                self._log(f"\n{sep}\n  🔍 Scanare: {total} fisiere .bin\n{sep}")

                scan_results = []; fisiere_goale = 0
                for idx, p in enumerate(self.selected_files, 1):
                    if cancelled(): abort(); return
                    sid, n_brut, n_lanes = quick_scan_bin(p)
                    scan_results.append((p, sid, n_brut, n_lanes))
                    fname = os.path.basename(p)
                    pct = int(2 + (idx / total) * 8)
                    if n_brut == 0:
                        self._log(f"  ⚠  [{sid}]  {fname}  →  fisier gol")
                        fisiere_goale += 1
                    else:
                        self._log(f"  ✓  [{sid}]  {fname}  →  ~{n_brut} inreg.  │  {n_lanes} benzi")
                    self._update_progress(pct, f"Scanare {idx}/{total}...")

                if cancelled(): abort(); return

                fisiere_valide = total - fisiere_goale
                goale_str = f"  ({fisiere_goale} goale)" if fisiere_goale else "  (toate valide)"
                self._log(f"\n  📋 Scanare: {fisiere_valide}/{total} valide{goale_str}\n")
                self._update_progress(10, "Se proceseaza datele .bin...")
                self._log("  🔄 Generare rapoarte Excel (.bin)...")

                rezultate_bin = process_multiple_files(
                    self.selected_files, stop_event=self.stop_event) or []
                if self.stop_event.is_set(): abort(); return

                if not rezultate_bin:
                    self._log("\n✗ Nicio data valida gasita in fisierele .bin.")
                else:
                    n_c = len(rezultate_bin)
                    self._update_progress(int(bin_pct_end * 0.7),
                                          f"Actualizez centralizatorul ({n_c} contoar{'e' if n_c!=1 else ''} .bin)...")
                    if cancelled(): abort(); return
                    update_centralizator_batch(
                        rezultate_bin, CENTRAL_FILE_FOLDER, log_callback=self._log)
                    self._update_progress(int(bin_pct_end * 0.98), "Centralizator .bin actualizat.")

                    t_ore  = sum(r["randuri"] for r in rezultate_bin)
                    t_b1   = sum(r["b1"] for r in rezultate_bin)
                    t_b2   = sum(r["b2"] for r in rezultate_bin)
                    lanes_i = ", ".join(f"{r['id']}:{r['n_lanes']}b" for r in rezultate_bin)
                    self._log(f"\n{sep}\n  📦 .BIN:")
                    self._log(f"     Contoar{'e' if n_c!=1 else ''}: {n_c}  │  Ore: {t_ore:,}  │  Trafic: {int(t_b1+t_b2):,}")
                    self._log(f"     Benzi: {lanes_i}\n{sep}")

                if cancelled(): abort(); return

            # ═══════════════════════════════════════════════════════════════
            # BLOC .LOG
            # ═══════════════════════════════════════════════════════════════
            if has_log:
                total_log = len(self.selected_log_files)
                self._update_progress(log_pct_start,
                    f"Se proceseaza {total_log} fisiere .log...")
                self._log(f"\n{sep}\n  🔍 Fisiere .log: {total_log}\n{sep}")

                for idx, fp in enumerate(self.selected_log_files, 1):
                    if cancelled(): abort(); return
                    self._log(f"  ✓  {os.path.basename(fp)}")
                    pct = int(log_pct_start + (idx / total_log) * 5)
                    self._update_progress(pct, f"Scanare .log {idx}/{total_log}...")

                if cancelled(): abort(); return

                self._update_progress(log_pct_start + 5,
                    "Se genereaza rapoartele Excel (.log)...")
                self._log("  🔄 Generare rapoarte Excel (.log)...")

                out_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)
                os.makedirs(out_dir, exist_ok=True)
                rezultate_log = process_log_files(
                    self.selected_log_files,
                    output_dir=out_dir,
                    stop_event=self.stop_event) or []

                if self.stop_event.is_set(): abort(); return

                if not rezultate_log:
                    self._log("\n✗ Nicio data valida gasita in fisierele .log.")
                else:
                    n_cl = len(rezultate_log)
                    self._update_progress(log_pct_start + 20,
                        f"Actualizez centralizatorul ({n_cl} contoar{'e' if n_cl!=1 else ''} .log)...")
                    self._log(f"\n  📊 Rapoarte .log generate: {n_cl} contoar{'e' if n_cl!=1 else ''}")

                    for r in rezultate_log:
                        self._log(f"  📥 Centralizator ← [{r['id']}]  "
                                  f"{r['randuri']} ore  │  "
                                  f"B1={r['b1']:,}  B2={r['b2']:,}")
                    if cancelled(): abort(); return
                    update_centralizator_batch(
                        rezultate_log, CENTRAL_FILE_FOLDER, log_callback=self._log)
                    self._update_progress(log_pct_start + 44, "Centralizator .log actualizat.")

                    t_ore_l = sum(r["randuri"] for r in rezultate_log)
                    t_veh_l = sum(r["b1"] + r["b2"] for r in rezultate_log)
                    lanes_l = ", ".join(f"{r['id']}:{r['n_lanes']}b" for r in rezultate_log)
                    self._log(f"\n{sep}\n  📦 .LOG:")
                    self._log(f"     Contoar{'e' if n_cl!=1 else ''}: {n_cl}  │  Ore: {t_ore_l:,}  │  Trafic: {t_veh_l:,}")
                    self._log(f"     Benzi: {lanes_l}\n{sep}")

            # ═══════════════════════════════════════════════════════════════
            # SUMAR FINAL
            # ═══════════════════════════════════════════════════════════════
            n_total = len(rezultate_bin) + len(rezultate_log)
            if n_total == 0:
                self._update_progress(0, "✗  Nicio data valida.")
                self.after(0, lambda: self.btn_process.configure(state="normal"))
                self.after(0, lambda: self.btn_cancel.configure(state="disabled"))
                return

            self._update_progress(100, f"✔  Finalizat — {n_total} contoar{'e' if n_total!=1 else ''} procesat{'e' if n_total!=1 else ''}.")

            rapoarte_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)

            lines = ["Procesare finalizata!\n"]
            if rezultate_bin:
                lines.append(f"📦 .BIN — {len(rezultate_bin)} contoar{'e' if n_total!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_bin):,} ore)")
            if rezultate_log:
                lines.append(f"📦 .LOG — {len(rezultate_log)} contoar{'e' if n_total!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_log):,} ore)")
            lines.append(f"\nRaport{'e' if n_total!=1 else ''} salvat{'e' if n_total!=1 else ''} in:\n{rapoarte_dir}")
            messagebox.showinfo("Succes", "\n".join(lines))

        except Exception as ex:
            self._log(f"\n✗ Eroare neasteptata: {ex}")
            self._update_progress(0, "✗  Eroare procesare.")
        finally:
            self._release_lock()  # eliberăm lock-ul indiferent de rezultat
            self.after(0, lambda: self.btn_process.configure(state="normal"))
            self.after(0, lambda: self.btn_cancel.configure(state="disabled"))
            self.after(0, lambda: setattr(self, "selected_files", []))
            self.after(0, lambda: setattr(self, "selected_log_files", []))
            self.after(0, lambda: self.lbl_files.configure(text="Niciun fișier .bin selectat."))
            self.after(0, lambda: self.lbl_log_files.configure(text="Niciun fișier .log selectat."))

    def _all_processing_background(self):
        """
        Thread background unificat: proceseaza .bin si/sau .log in ordine.
        """
        sep = "─" * 52

        def cancelled():
            return self.stop_event.is_set()

        def abort(msg="⛔ Procesare anulata de utilizator."):
            self._log(f"\n{msg}")
            self._update_progress(0, "⛔ Anulat.")
            self.after(0, lambda: self.btn_process.configure(state="normal"))
            self.after(0, lambda: self.btn_cancel.configure(state="disabled"))

        try:
            _t_run_start = time.perf_counter()
            _perf_phases = []
            _t_phase = _t_run_start
            _n_bin_ok = 0
            _n_log_ok = 0

            has_bin = bool(self.selected_files)
            has_log = bool(self.selected_log_files)

            bin_pct_end   = 95 if not has_log else 60
            log_pct_start = bin_pct_end if has_bin else 5

            rezultate_bin = []
            rezultate_log = []

            # ═══════════════════════════════════════════════════════════════
            # BLOC .BIN
            # ═══════════════════════════════════════════════════════════════
            if has_bin:
                total = len(self.selected_files)
                self._update_progress(2, f"Se scaneaza {total} fisiere .bin...")
                self._log(f"\n{sep}\n  🔍 Scanare: {total} fisiere .bin\n{sep}")

                fisiere_goale = 0
                for idx, p in enumerate(self.selected_files, 1):
                    if cancelled(): abort(); return
                    sid, n_brut, n_lanes = quick_scan_bin(p)
                    fname = os.path.basename(p)
                    pct = int(2 + (idx / total) * 8)
                    if n_brut == 0:
                        self._log(f"  ⚠  [{sid}]  {fname}  →  fisier gol")
                        fisiere_goale += 1
                    else:
                        self._log(f"  ✓  [{sid}]  {fname}  →  ~{n_brut} inreg.  │  {n_lanes} benzi")
                    self._update_progress(pct, f"Scanare {idx}/{total}...")

                if cancelled(): abort(); return

                fisiere_valide = total - fisiere_goale
                goale_str = f"  ({fisiere_goale} goale)" if fisiere_goale else "  (toate valide)"
                self._log(f"\n  📋 Scanare: {fisiere_valide}/{total} valide{goale_str}\n")
                self._update_progress(10, "Se proceseaza datele .bin...")
                self._log("  🔄 Generare rapoarte Excel (.bin)...")

                _now = time.perf_counter()
                _perf_phases.append(("Scanare .bin", _now - _t_phase))
                _t_phase = _now

                # ── Callback progres per contor (apelat din thread-ul bin_parser) ─
                _bin_proc_start = 10
                _bin_proc_end   = int(bin_pct_end * 0.65)

                def _bin_progress(site_id, n_ore, idx_ct, total_ct):
                    if cancelled(): return
                    pct = int(_bin_proc_start +
                              (idx_ct / total_ct) * (_bin_proc_end - _bin_proc_start))
                    self.after(0, lambda p=pct, s=site_id, n=n_ore, i=idx_ct, t=total_ct:
                        self._update_progress(
                            p, f"💾 SQLite+Excel [{s}]  {n:,} ore  ({i}/{t})"))
                    self.after(0, lambda s=site_id, n=n_ore:
                        self._log(f"  💾 SQLite ← [{s}]  {n:,} rânduri orare"))

                reset_ddp_perf()   # ⏱ instrumentare foaie "Date detaliate prelucrate"
                rezultate_bin = process_multiple_files(
                    self.selected_files,
                    stop_event=self.stop_event,
                    progress_callback=_bin_progress) or []
                if self.stop_event.is_set(): abort(); return

                _now = time.perf_counter()
                _perf_phases.append(("Procesare + SQLite + Excel .bin", _now - _t_phase))
                _perf_phases.append(("  └─ din care: foaia Date detaliate prelucrate",
                                      get_ddp_perf_seconds()))
                _t_phase = _now

                if not rezultate_bin:
                    self._log("\n✗ Nicio data valida gasita in fisierele .bin.")
                else:
                    n_c = len(rezultate_bin)
                    _n_bin_ok = n_c
                    self._update_progress(int(bin_pct_end * 0.65),
                        f"Actualizez centralizatorul ({n_c} contoar{'e' if n_c!=1 else ''} .bin)...")
                    if cancelled(): abort(); return
                    update_centralizator_batch(
                        rezultate_bin, CENTRAL_FILE_FOLDER, log_callback=self._log)
                    self._update_progress(int(bin_pct_end * 0.95), "Centralizator .bin actualizat.")

                    t_ore  = sum(r["randuri"] for r in rezultate_bin)
                    t_b1   = sum(r["b1"] for r in rezultate_bin)
                    t_b2   = sum(r["b2"] for r in rezultate_bin)
                    lanes_i = ", ".join(f"{r['id']}:{r['n_lanes']}b" for r in rezultate_bin)
                    self._log(f"\n{sep}\n  📦 .BIN:")
                    self._log(f"     Contoar{'e' if n_c!=1 else ''}: {n_c}  │  Ore: {t_ore:,}  │  Trafic: {int(t_b1+t_b2):,}")
                    self._log(f"     Benzi: {lanes_i}\n{sep}")

                _now = time.perf_counter()
                _perf_phases.append(("Centralizator .bin", _now - _t_phase))
                _t_phase = _now

                if cancelled(): abort(); return

            # ═══════════════════════════════════════════════════════════════
            # BLOC .LOG
            # ═══════════════════════════════════════════════════════════════
            if has_log:
                total_log = len(self.selected_log_files)
                self._update_progress(log_pct_start,
                    f"Se proceseaza {total_log} fisiere .log...")
                self._log(f"\n{sep}\n  🔍 Fisiere .log: {total_log}\n{sep}")

                for idx, fp in enumerate(self.selected_log_files, 1):
                    if cancelled(): abort(); return
                    self._log(f"  ✓  {os.path.basename(fp)}")
                    pct = int(log_pct_start + (idx / total_log) * 5)
                    self._update_progress(pct, f"Scanare .log {idx}/{total_log}...")

                if cancelled(): abort(); return

                _now = time.perf_counter()
                _perf_phases.append(("Scanare .log", _now - _t_phase))
                _t_phase = _now

                self._update_progress(log_pct_start + 5,
                    "Se genereaza rapoartele Excel (.log)...")
                self._log("  🔄 Generare rapoarte Excel (.log)...")

                out_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)
                os.makedirs(out_dir, exist_ok=True)

                # ── Callback progres per contor .log ─────────────────────
                _log_proc_start = log_pct_start + 5
                _log_proc_end   = log_pct_start + 18

                def _log_progress(site_id, n_ore, idx_ct, total_ct):
                    if cancelled(): return
                    pct = int(_log_proc_start +
                              (idx_ct / total_ct) * (_log_proc_end - _log_proc_start))
                    self.after(0, lambda p=pct, s=site_id, n=n_ore, i=idx_ct, t=total_ct:
                        self._update_progress(
                            p, f"💾 SQLite+Excel [{s}]  {n:,} ore VEK  ({i}/{t})"))
                    self.after(0, lambda s=site_id, n=n_ore:
                        self._log(f"  💾 SQLite ← [{s}]  {n:,} rânduri orare (VEK)"))

                reset_ddp_perf()   # ⏱ instrumentare foaie "Date detaliate prelucrate"
                rezultate_log = process_log_files(
                    self.selected_log_files,
                    output_dir=out_dir,
                    stop_event=self.stop_event,
                    progress_callback=_log_progress) or []

                if self.stop_event.is_set(): abort(); return

                _now = time.perf_counter()
                _perf_phases.append(("Procesare + SQLite + Excel .log", _now - _t_phase))
                _perf_phases.append(("  └─ din care: foaia Date detaliate prelucrate",
                                      get_ddp_perf_seconds()))
                _t_phase = _now

                if not rezultate_log:
                    self._log("\n✗ Nicio data valida gasita in fisierele .log.")
                else:
                    n_cl = len(rezultate_log)
                    _n_log_ok = n_cl
                    self._update_progress(log_pct_start + 20,
                        f"Actualizez centralizatorul ({n_cl} contoar{'e' if n_cl!=1 else ''} .log)...")
                    self._log(f"\n  📊 Rapoarte .log generate: {n_cl} contoar{'e' if n_cl!=1 else ''}")

                    for r in rezultate_log:
                        self._log(f"  📥 Centralizator ← [{r['id']}]  "
                                  f"{r['randuri']} ore  │  "
                                  f"B1={r['b1']:,}  B2={r['b2']:,}")
                    if cancelled(): abort(); return
                    update_centralizator_batch(
                        rezultate_log, CENTRAL_FILE_FOLDER, log_callback=self._log)
                    self._update_progress(log_pct_start + 44, "Centralizator .log actualizat.")

                    t_ore_l = sum(r["randuri"] for r in rezultate_log)
                    t_veh_l = sum(r["b1"] + r["b2"] for r in rezultate_log)
                    lanes_l = ", ".join(f"{r['id']}:{r['n_lanes']}b" for r in rezultate_log)
                    self._log(f"\n{sep}\n  📦 .LOG:")
                    self._log(f"     Contoar{'e' if n_cl!=1 else ''}: {n_cl}  │  Ore: {t_ore_l:,}  │  Trafic: {t_veh_l:,}")
                    self._log(f"     Benzi: {lanes_l}\n{sep}")

            if has_log:
                _now = time.perf_counter()
                _perf_phases.append(("Centralizator .log", _now - _t_phase))
                _t_phase = _now

            # ═══════════════════════════════════════════════════════════════
            # SUMAR FINAL
            # ═══════════════════════════════════════════════════════════════
            n_total = len(rezultate_bin) + len(rezultate_log)
            if n_total == 0:
                self._update_progress(0, "✗  Nicio data valida.")
                self.after(0, lambda: self.btn_process.configure(state="normal"))
                self.after(0, lambda: self.btn_cancel.configure(state="disabled"))
                return

            self._update_progress(100,
                f"✔  Finalizat — {n_total} contoar{'e' if n_total!=1 else ''} procesat{'e' if n_total!=1 else ''}.")

            rapoarte_dir = os.path.join(CENTRAL_FILE_FOLDER, RAPOARTE_PEEK_FOLDER)

            _t_run_elapsed = time.perf_counter() - _t_run_start
            self._log(f"\n{sep}")
            self._log(f"  ⏱ RAPORT DE PERFORMANȚĂ (unde s-a dus timpul):")
            self._log(_format_perf_summary(_perf_phases, _t_run_elapsed))

            _perf_dict = dict(_perf_phases)
            _avg_lines = []
            if _n_bin_ok:
                _t_proc_bin = _perf_dict.get("Procesare + SQLite + Excel .bin", 0.0)
                _avg_lines.append(
                    f"     Timp mediu / post .bin   ({_n_bin_ok:>2} contoare):   {_t_proc_bin / _n_bin_ok:6.1f}s")
            if _n_log_ok:
                _t_proc_log = _perf_dict.get("Procesare + SQLite + Excel .log", 0.0)
                _avg_lines.append(
                    f"     Timp mediu / post .log   ({_n_log_ok:>2} contoare):   {_t_proc_log / _n_log_ok:6.1f}s")
            if (_n_bin_ok + _n_log_ok):
                _avg_lines.append(
                    f"     Timp mediu / post total  ({_n_bin_ok + _n_log_ok:>2} contoare):   {_t_run_elapsed / (_n_bin_ok + _n_log_ok):6.1f}s")
            if _avg_lines:
                self._log("\n".join(_avg_lines))
            self._log(sep)

            lines = ["Procesare finalizata!\n"]
            if rezultate_bin:
                lines.append(f"📦 .BIN — {len(rezultate_bin)} contoar{'e' if len(rezultate_bin)!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_bin):,} ore)")
            if rezultate_log:
                lines.append(f"📦 .LOG — {len(rezultate_log)} contoar{'e' if len(rezultate_log)!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_log):,} ore)")
            lines.append(f"\nRaport{'e' if n_total!=1 else ''} salvat{'e' if n_total!=1 else ''} in:\n{rapoarte_dir}")
            self.after(0, lambda: messagebox.showinfo("Succes", "\n".join(lines)))

        except Exception as ex:
            self._log(f"\n✗ Eroare neasteptata: {ex}")
            self._update_progress(0, "✗  Eroare procesare.")
        finally:
            self._release_lock()
            self.after(0, lambda: self.btn_process.configure(state="normal"))
            self.after(0, lambda: self.btn_cancel.configure(state="disabled"))
            self.after(0, lambda: setattr(self, "selected_files", []))
            self.after(0, lambda: setattr(self, "selected_log_files", []))
            self.after(0, lambda: self.lbl_files.configure(text="Niciun fișier .bin selectat."))
            self.after(0, lambda: self.lbl_log_files.configure(text="Niciun fișier .log selectat."))

    def _cancel_processing(self):
        self.stop_event.set()
        self.lbl_status.configure(text="Se anulează procesarea…")
        self._log("⚠ Anulare solicitată de utilizator.")

    def _force_release_lock(self):
        """Eliberează manual lock-ul de procesare — pentru situații de urgență."""
        lock_path = self._get_lock_path()
        if not lock_path or not os.path.exists(lock_path):
            messagebox.showinfo("Lock", "Nu există niciun lock activ de procesare.",
                                parent=self)
            return
        try:
            info = open(lock_path, encoding="utf-8").read().strip()
        except Exception:
            info = "necunoscut"
        if messagebox.askyesno(
            "Eliberează lock",
            f"Există un lock activ de la:\n\n  {info}\n\n"
            f"Ești sigur că vrei să îl eliberezi forțat?\n"
            f"(Folosește doar dacă procesarea s-a terminat sau aplicația a căzut)",
            parent=self
        ):
            self._release_lock()
            self._log("🔓  Lock de procesare eliberat manual.")
            messagebox.showinfo("Gata", "Lock-ul a fost eliberat. Poți procesa acum.",
                                parent=self)

    # ── Selectare și procesare fișiere .log ───────────────────────────────────
    def _choose_log_files(self):
        paths = filedialog.askopenfilenames(
            title="Selectează fișiere .log VEK",
            filetypes=[("Fișiere LOG", "*.log"), ("Toate fișierele", "*.*")],
        )
        if not paths:
            return
        self.selected_log_files = list(paths)
        n = len(paths)
        def _sid(p):
            base = os.path.splitext(os.path.basename(p))[0]
            return base.split("_")[0]
        ct = sorted(set(_sid(p) for p in paths))
        nc = len(ct)
        self.lbl_log_files.configure(
            text=f"{n} fișier{'e' if n != 1 else ''} .log selectat{'e' if n != 1 else ''} "
                 f"de la {nc} contoar{'e' if nc != 1 else ''}.")
        sep = "─" * 52
        self._log(f"\n{sep}")
        self._log(f"  📂 Fișier{'e' if n != 1 else ''} .log: {n}  │  Contoar{'e' if nc != 1 else ''}: {nc}")
        if nc <= 20: self._log(f"  📋 {', '.join(ct)}")
        else:        self._log(f"  📋 {', '.join(ct[:20])} ... (+{nc - 20})")
        self._log(sep)

    # ── Gestionare Contoare ───────────────────────────────────────────────────
    def _open_contoare_manager(self):
        if CTK_AVAILABLE:
            win = ctk.CTkToplevel(self)
        else:
            win = tk.Toplevel(self)

        win.title("Gestionare Contoare")
        win.resizable(True, True)
        win.minsize(1000, 460)
        W, H = 1150, 580
        sw = win.winfo_screenwidth(); sh = win.winfo_screenheight()
        win.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
        win.grab_set()
        win.focus_force()

        # Header cu logo + titlu
        _add_logo_header(win, "Gestionare Contoare", app_ref=self)
        _ctk_label(win,
                   f"Fișier: {os.path.join(CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME)}",
                   font=FONT_SMALL, text_color="#999999").pack(pady=(0, 4))

        # Treeview cu stil
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=16, pady=4)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Peek.Treeview",
                        background="#FFFFFF", fieldbackground="#FFFFFF",
                        font=("Segoe UI", 10), rowheight=26,
                        borderwidth=0, relief="flat")
        style.configure("Peek.Treeview.Heading",
                        font=("Segoe UI", 10, "bold"),
                        background="#1A2533", foreground="white",
                        relief="flat", borderwidth=0)
        style.map("Peek.Treeview",
                  background=[("selected", "#2471A3")],
                  foreground=[("selected", "white")])

        cols = ("Contor", "Drum", "Poziție km", "Localitate", "Tip", "IP",
                "DRDP", "Lat", "Lng")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                            height=14, style="Peek.Treeview")
        cw = {"Contor": 70, "Drum": 85, "Poziție km": 80,
              "Localitate": 130, "Tip": 160, "IP": 115,
              "DRDP": 90, "Lat": 95, "Lng": 95}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=cw.get(c, 90), anchor="center")

        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=tree.yview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        if CTK_AVAILABLE:
            ctk.CTkFrame(win, height=1, fg_color="#CCCCCC").pack(fill="x", padx=16, pady=(6, 4))
        else:
            ttk.Separator(win, orient="horizontal").pack(fill="x", padx=16, pady=(6, 4))

        # Butoane scalabile
        btn_outer = ctk.CTkFrame(win, fg_color="transparent") if CTK_AVAILABLE else tk.Frame(win)
        btn_outer.pack(fill="x", padx=16, pady=(0, 14))
        for col in range(4):
            btn_outer.columnconfigure(col, weight=1, uniform="mgr")

        def _get_contor_data_merged(ct_id):
            """
            Returnează datele unui contor îmbinate din ambele surse:
            SQLite contoare.db (primar, are DRDP/lat/lng) + Excel (fallback).
            """
            data = {}
            # Întâi Excel (date de bază vechi)
            db_excel = _load_contoare_db()
            if ct_id in db_excel:
                data.update(db_excel[ct_id])
            # Suprascrie cu datele din SQLite (mai recente, inclusiv DRDP/lat/lng)
            try:
                from database import get_contoare_db
                d_sql = get_contoare_db().get(ct_id)
                if d_sql:
                    data.update(d_sql)
            except Exception:
                pass
            return data

        def refresh_tree():
            for row in tree.get_children():
                tree.delete(row)
            # Citim din SQLite — sursa principală (are DRDP/lat/lng)
            try:
                from database import get_contoare_db, CONTOARE_DB
                db_sql = get_contoare_db().get_all()
            except Exception:
                db_sql = {}
            # Fallback: completăm cu contoare din Excel care nu sunt încă în SQLite
            db_excel = _load_contoare_db()

            # Excludem codurile vechi redenumite (din contor_alias)
            try:
                import sqlite3 as _sqlite3
                _conn = _sqlite3.connect(CONTOARE_DB, timeout=5)
                _coduri_vechi = {
                    r[0] for r in
                    _conn.execute("SELECT cod_vechi FROM contor_alias").fetchall()
                }
                _conn.close()
            except Exception:
                _coduri_vechi = set()

            ct_ids = sorted(
                (set(db_sql.keys()) | set(db_excel.keys())) - _coduri_vechi
            )

            def _fmt_coord(v):
                try:
                    return f"{float(v):.6f}" if v not in ("", None) else ""
                except Exception:
                    return str(v) if v else ""

            for ct in ct_ids:
                d = db_sql.get(ct) or db_excel.get(ct) or {}
                tree.insert("", "end", values=(
                    ct,
                    d.get("Drum","") or d.get("drum",""),
                    d.get("Pozitie_km","") or d.get("pozitie_km",""),
                    d.get("Localitate","") or d.get("localitate",""),
                    d.get("Tip","") or d.get("tip",""),
                    d.get("IP","") or d.get("ip",""),
                    d.get("DRDP","") or d.get("drdp",""),
                    _fmt_coord(d.get("lat") or d.get("Lat") or ""),
                    _fmt_coord(d.get("lng") or d.get("Lng") or ""),
                ))

        def on_adauga():
            _open_contor_dialog(win, mode="add", on_save=refresh_tree, app_ref=self)

        def on_editeaza():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Atenție", "Selectează un contoar!", parent=win)
                return
            vals = tree.item(sel[0])["values"]
            ct_id = str(vals[0])
            _open_contor_dialog(win, mode="edit", ct_id=ct_id,
                                data=_get_contor_data_merged(ct_id),
                                on_save=refresh_tree, app_ref=self)

        def on_sterge():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Atenție", "Selectează un contoar!", parent=win)
                return
            vals = tree.item(sel[0])["values"]
            ct_id = str(vals[0])
            msg = f"Ești sigur că vrei să ștergi contoarul {ct_id}?"
            if messagebox.askyesno("Confirmare ștergere", msg, parent=win):
                db = _load_contoare_db()
                if ct_id in db:
                    del db[ct_id]
                    _save_contoare_db(db)
                    _delete_contor_from_centralizator(ct_id)
                    # Ștergere și din contoare.db SQLite
                    try:
                        from database import get_contoare_db
                        get_contoare_db().delete(ct_id)
                    except Exception as _e:
                        print(f"[WARN] contoare.db delete eroare [{ct_id}]: {_e}")
                    self._log(f"  🗑  Contoar {ct_id} șters din Gestionare și din Centralizator.")
                    refresh_tree()

        for col, (txt, cmd, ck) in enumerate([
            ("➕  Adaugă",   on_adauga,   "success"),
            ("✏️  Editează", on_editeaza, "primary"),
            ("🗑️  Șterge",  on_sterge,   "danger"),
            ("✖  Închide",  win.destroy, "secondary"),
        ]):
            b = _ctk_btn(btn_outer, txt, cmd, ck)
            b.grid(row=0, column=col, padx=5, pady=6, sticky="ew")

        refresh_tree()
        tree.bind("<Double-1>", lambda e: on_editeaza())

    # ── Procesare manuală MZL ─────────────────────────────────────────────────
    def _open_mzl_manual_dialog(self):
        """
        Fereastră pentru editarea manuală a MZL (Media Zilnică Lunară).

        Flux:
          1. Operator scrie numărul contorului
          2. Aplicația interogează DB → afișează localitate + ani/luni disponibile
          3. Operator selectează an și lună
          4. Aplicația afișează valoarea curentă calculată + cea manuală existentă
          5. Operator introduce noua valoare MZL
          6. Confirmare: "Modifici MZL de la X la Y?" → salvare în DB cu username
        """
        try:
            from database import get_traffic_db, get_contoare_db
        except ImportError:
            messagebox.showerror(
                "Modul lipsă",
                "Modulul database.py nu este disponibil.\n"
                "Asigurați-vă că fișierul database.py există în același folder.",
            )
            return

        import getpass

        tdb = get_traffic_db()
        cdb = get_contoare_db()

        # ── Creare fereastră — singleton: dacă există deja, o aducem în față ──
        if hasattr(self, '_mzl_win') and self._mzl_win is not None:
            try:
                if self._mzl_win.winfo_exists():
                    self._mzl_win.lift()
                    self._mzl_win.focus_force()
                    return
            except Exception:
                pass
            self._mzl_win = None

        if CTK_AVAILABLE:
            win = ctk.CTkToplevel(self)
        else:
            win = tk.Toplevel(self)

        self._mzl_win = win

        def _on_close():
            try:
                win.grab_release()
            except Exception:
                pass
            self._mzl_win = None
            try:
                win.destroy()
            except Exception:
                pass

        win.protocol("WM_DELETE_WINDOW", _on_close)
        win.title("Procesare manuală MZL")
        win.resizable(False, False)
        # Aducem fereastra în față — lift + focus după randare completă
        def _bring_to_front():
            if not win.winfo_exists():
                return
            win.lift()
            win.attributes("-topmost", True)
            win.after(200, lambda: win.attributes("-topmost", False) if win.winfo_exists() else None)
            win.focus_force()
            # grab_local în loc de grab_set — nu blochează alte instanțe ale
            # aplicației pe alte PC-uri care accesează aceeași bază de date
            try:
                win.grab_set_global() if False else win.grab_set()
            except Exception:
                pass
        win.after(150, _bring_to_front)

        W, H = 520, 560
        sw = win.winfo_screenwidth(); sh = win.winfo_screenheight()
        win.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        _add_logo_header(win, "Procesare manuală MZL", app_ref=self)

        # ── Stare internă ─────────────────────────────────────────────────────
        var_contor    = tk.StringVar()
        var_an        = tk.StringVar()
        var_luna      = tk.StringVar()
        var_mzl_nou   = tk.StringVar()
        var_operator  = tk.StringVar(value=getpass.getuser())

        # Valori curente afișate (non-editabile)
        var_localitate   = tk.StringVar(value="—")
        var_mzl_calc     = tk.StringVar(value="—")
        var_mzl_manual   = tk.StringVar(value="—")

        LUNI_NUME = {
            1: "Ianuarie", 2: "Februarie", 3: "Martie",    4: "Aprilie",
            5: "Mai",      6: "Iunie",     7: "Iulie",      8: "August",
            9: "Septembrie", 10: "Octombrie", 11: "Noiembrie", 12: "Decembrie",
        }
        LUNI_INV = {v: k for k, v in LUNI_NUME.items()}

        # ── Layout principal ──────────────────────────────────────────────────
        pad = {"padx": 28, "pady": 5}

        def mk_row(parent, lbl_text, row):
            """Rând label + widget, returnează frame dreapta."""
            if CTK_AVAILABLE:
                ctk.CTkLabel(parent, text=lbl_text, font=FONT_SMALL,
                             anchor="e", width=130).grid(
                    row=row, column=0, padx=(0, 10), pady=6, sticky="e")
            else:
                tk.Label(parent, text=lbl_text, font=FONT_SMALL,
                         anchor="e", width=16).grid(
                    row=row, column=0, padx=(0, 10), pady=6, sticky="e")

        if CTK_AVAILABLE:
            body = ctk.CTkFrame(win, fg_color="transparent")
        else:
            body = tk.Frame(win)
        body.pack(fill="x", padx=20, pady=4)
        body.columnconfigure(1, weight=1)

        # ── Rând 0: Număr contor ──────────────────────────────────────────────
        mk_row(body, "Număr contor *", 0)
        frm_ct = ctk.CTkFrame(body, fg_color="transparent") if CTK_AVAILABLE \
            else tk.Frame(body)
        frm_ct.grid(row=0, column=1, sticky="w", pady=6)

        entry_contor = _ctk_entry(frm_ct, textvariable=var_contor, width=120)
        entry_contor.pack(side="left", padx=(0, 8))

        def on_cauta_contor():
            ct = var_contor.get().strip()
            if not ct:
                return
            btn_cauta.configure(state="disabled")
            var_localitate.set("Se caută...")

            def _do_cauta():
                try:
                    loc = cdb.get_localitate(ct) if ct else ""
                    ani = tdb.get_ani_disponibili(ct)
                except Exception as e:
                    win.after(0, lambda: var_localitate.set("Eroare DB"))
                    win.after(0, lambda: btn_cauta.configure(state="normal"))
                    win.after(0, lambda err=str(e): messagebox.showerror(
                        "Eroare DB", err, parent=win))
                    return

                def _update_ui():
                    btn_cauta.configure(state="normal")
                    var_localitate.set(loc if loc else "Necunoscută")
                    if not ani:
                        messagebox.showwarning("Contor negăsit",
                            f"Contorul {ct} nu are date în baza de date.",
                            parent=win)
                        var_localitate.set("—")
                        return
                    combo_an.configure(values=[str(a) for a in ani])
                    var_an.set(str(ani[-1]))
                    on_an_changed()

                win.after(0, _update_ui)

            threading.Thread(target=_do_cauta, daemon=True).start()

        btn_cauta = _ctk_btn(frm_ct, "🔍 Caută", on_cauta_contor,
                             "primary", width=90)
        btn_cauta.pack(side="left")

        # ── Rând 1: Localitate (read-only) ───────────────────────────────────
        mk_row(body, "Localitate", 1)
        if CTK_AVAILABLE:
            ctk.CTkLabel(body, textvariable=var_localitate,
                         font=("Segoe UI", 11, "bold"),
                         text_color="#1A5276",
                         anchor="w").grid(row=1, column=1, sticky="w", pady=4)
        else:
            tk.Label(body, textvariable=var_localitate,
                     font=("Segoe UI", 11, "bold"),
                     fg="#1A5276", anchor="w").grid(
                row=1, column=1, sticky="w", pady=4)

        # ── Rând 2: An ────────────────────────────────────────────────────────
        mk_row(body, "An", 2)
        combo_an = (ctk.CTkComboBox if CTK_AVAILABLE else ttk.Combobox)(
            body, variable=var_an,
            values=[], width=110,
            state="readonly",
            **({} if not CTK_AVAILABLE else {"height": 30, "corner_radius": 6})
        )
        combo_an.grid(row=2, column=1, sticky="w", pady=6)

        def on_an_changed(*_):
            ct = var_contor.get().strip()
            an_str = var_an.get()
            if not ct or not an_str:
                return
            try:
                an = int(an_str)
            except ValueError:
                return
            luni = tdb.get_luni_disponibile(ct, an)
            combo_luna.configure(values=[LUNI_NUME[l] for l in luni if l in LUNI_NUME])
            if luni:
                var_luna.set(LUNI_NUME[luni[-1]])
                on_luna_changed()

        combo_an.bind("<<ComboboxSelected>>", on_an_changed)
        if CTK_AVAILABLE:
            combo_an.configure(command=on_an_changed)

        # ── Rând 3: Lună ──────────────────────────────────────────────────────
        mk_row(body, "Lună", 3)
        combo_luna = (ctk.CTkComboBox if CTK_AVAILABLE else ttk.Combobox)(
            body, variable=var_luna,
            values=[], width=160,
            state="readonly",
            **({} if not CTK_AVAILABLE else {"height": 30, "corner_radius": 6})
        )
        combo_luna.grid(row=3, column=1, sticky="w", pady=6)

        def on_luna_changed(*_):
            """Afișează MZL calculat și cel manual existent pentru an+lună.
            Rulează query-urile DB în thread separat — nu blochează GUI-ul."""
            ct     = var_contor.get().strip()
            an_str = var_an.get()
            luna_n = var_luna.get()
            if not ct or not an_str or not luna_n:
                return
            try:
                an   = int(an_str)
                luna = LUNI_INV.get(luna_n)
                if luna is None:
                    return
            except ValueError:
                return

            var_mzl_calc.set("Se calculează...")
            var_mzl_manual.set("...")

            def _do_calcul():
                import calendar
                try:
                    rows_zi = tdb._conn().execute("""
                        SELECT zi, SUM(total_general) AS total_zi, COUNT(*) AS ore_zi
                        FROM inregistrari_orare
                        WHERE contor=? AND an=? AND luna=?
                        GROUP BY zi
                    """, (ct, an, luna)).fetchall()

                    zile_luna   = calendar.monthrange(an, luna)[1]
                    zile_valide = [r for r in rows_zi if int(r["ore_zi"] or 0) >= 22]
                    n_zile_val  = len(zile_valide)

                    if n_zile_val > 0:
                        total_zile_valide = sum(int(r["total_zi"] or 0) for r in zile_valide)
                        mzl_auto = round(total_zile_valide / n_zile_val)
                        mzl_calc_str = (f"{mzl_auto:,}  "
                                        f"({n_zile_val}/{zile_luna} zile valide)")
                    else:
                        mzl_calc_str = "Fără date valide (< 22 ore/zi)"
                        mzl_auto = None
                except Exception as e:
                    mzl_calc_str = f"Eroare: {e}"
                    mzl_auto = None

                try:
                    manuale  = tdb.get_mzl_manual(ct)
                    val_man  = manuale.get((an, luna))
                except Exception:
                    val_man = None

                def _update():
                    var_mzl_calc.set(mzl_calc_str)
                    if val_man is not None:
                        var_mzl_manual.set(f"{int(val_man):,}")
                        var_mzl_nou.set(str(int(val_man)))
                    else:
                        var_mzl_manual.set("(nesetat)")
                        var_mzl_nou.set("")

                win.after(0, _update)

            threading.Thread(target=_do_calcul, daemon=True).start()

        combo_luna.bind("<<ComboboxSelected>>", on_luna_changed)
        if CTK_AVAILABLE:
            combo_luna.configure(command=on_luna_changed)

        # ── Separator ─────────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            ctk.CTkFrame(win, height=1, fg_color="#DDDDDD").pack(
                fill="x", padx=20, pady=(8, 2))
        else:
            ttk.Separator(win, orient="horizontal").pack(
                fill="x", padx=20, pady=(8, 2))

        # ── Card valori curente ───────────────────────────────────────────────
        if CTK_AVAILABLE:
            card_val = ctk.CTkFrame(win, fg_color="#F0F4F8", corner_radius=8)
        else:
            card_val = tk.Frame(win, bg="#F0F4F8", relief="flat", bd=1)
        card_val.pack(fill="x", padx=20, pady=6)
        card_val.columnconfigure(1, weight=1)

        def info_row(parent, label, var, row, bold=False, color="#1A1A2E"):
            if CTK_AVAILABLE:
                ctk.CTkLabel(parent, text=label, font=FONT_SMALL,
                             anchor="e", width=160,
                             text_color="#666666").grid(
                    row=row, column=0, padx=(12, 6), pady=4, sticky="e")
                ctk.CTkLabel(parent, textvariable=var,
                             font=("Segoe UI", 11, "bold") if bold else FONT_SMALL,
                             text_color=color, anchor="w").grid(
                    row=row, column=1, padx=(0, 12), pady=4, sticky="w")
            else:
                tk.Label(parent, text=label, font=FONT_SMALL,
                         fg="#666666", bg="#F0F4F8",
                         anchor="e", width=20).grid(
                    row=row, column=0, padx=(12, 6), pady=4, sticky="e")
                tk.Label(parent, textvariable=var,
                         font=("Segoe UI", 10, "bold") if bold else FONT_SMALL,
                         fg=color, bg="#F0F4F8", anchor="w").grid(
                    row=row, column=1, padx=(0, 12), pady=4, sticky="w")

        info_row(card_val, "MZL calculat automat:", var_mzl_calc,  0)
        info_row(card_val, "MZL manual curent:",    var_mzl_manual, 1,
                 bold=True, color="#1A5276")

        # ── Separator ─────────────────────────────────────────────────────────
        if CTK_AVAILABLE:
            ctk.CTkFrame(win, height=1, fg_color="#DDDDDD").pack(
                fill="x", padx=20, pady=(8, 2))
        else:
            ttk.Separator(win, orient="horizontal").pack(
                fill="x", padx=20, pady=(8, 2))

        # ── Câmpuri editabile ─────────────────────────────────────────────────
        if CTK_AVAILABLE:
            edit_frame = ctk.CTkFrame(win, fg_color="transparent")
        else:
            edit_frame = tk.Frame(win)
        edit_frame.pack(fill="x", padx=20, pady=4)
        edit_frame.columnconfigure(1, weight=1)

        def mk_edit_row(parent, lbl, var, row, placeholder=""):
            if CTK_AVAILABLE:
                ctk.CTkLabel(parent, text=lbl, font=FONT_SMALL,
                             anchor="e", width=130).grid(
                    row=row, column=0, padx=(0, 10), pady=6, sticky="e")
                e = ctk.CTkEntry(parent, textvariable=var, width=200,
                                 height=32, corner_radius=6,
                                 placeholder_text=placeholder)
            else:
                tk.Label(parent, text=lbl, font=FONT_SMALL,
                         anchor="e", width=16).grid(
                    row=row, column=0, padx=(0, 10), pady=6, sticky="e")
                e = ttk.Entry(parent, textvariable=var, width=24)
            e.grid(row=row, column=1, sticky="w", pady=6)
            return e

        mk_edit_row(edit_frame, "MZL nou *", var_mzl_nou, 0,
                    placeholder="ex: 1250")

        # ── Operator — read-only, preluat din Windows (getpass.getuser()) ─────
        if CTK_AVAILABLE:
            ctk.CTkLabel(edit_frame, text="Operator", font=FONT_SMALL,
                         anchor="e", width=130).grid(
                row=1, column=0, padx=(0, 10), pady=6, sticky="e")
            ctk.CTkLabel(edit_frame, textvariable=var_operator,
                         font=("Segoe UI", 11, "bold"),
                         text_color="#1A5276", anchor="w").grid(
                row=1, column=1, sticky="w", pady=6)
        else:
            tk.Label(edit_frame, text="Operator", font=FONT_SMALL,
                     anchor="e", width=16).grid(
                row=1, column=0, padx=(0, 10), pady=6, sticky="e")
            tk.Label(edit_frame, textvariable=var_operator,
                     font=("Segoe UI", 10, "bold"),
                     fg="#1A5276", anchor="w").grid(
                row=1, column=1, sticky="w", pady=6)

        # ── Notă operator ─────────────────────────────────────────────────────
        mk_row(edit_frame, "Observații", 2)
        var_obs = tk.StringVar()
        if CTK_AVAILABLE:
            ctk.CTkEntry(edit_frame, textvariable=var_obs, width=280,
                         height=32, corner_radius=6,
                         placeholder_text="(opțional)").grid(
                row=2, column=1, sticky="w", pady=6)
        else:
            ttk.Entry(edit_frame, textvariable=var_obs, width=34).grid(
                row=2, column=1, sticky="w", pady=6)

        # ── Separator + butoane ───────────────────────────────────────────────
        if CTK_AVAILABLE:
            ctk.CTkFrame(win, height=1, fg_color="#CCCCCC").pack(
                fill="x", padx=20, pady=(10, 4))
        else:
            ttk.Separator(win, orient="horizontal").pack(
                fill="x", padx=20, pady=(10, 4))

        btn_row = ctk.CTkFrame(win, fg_color="transparent") if CTK_AVAILABLE \
            else tk.Frame(win)
        btn_row.pack(pady=10)

        def on_salveaza():
            ct      = var_contor.get().strip()
            an_str  = var_an.get().strip()
            luna_n  = var_luna.get().strip()
            mzl_str = var_mzl_nou.get().strip()
            op      = var_operator.get().strip()
            obs     = var_obs.get().strip()

            # Validări
            if not ct:
                messagebox.showwarning("Câmp lipsă",
                    "Introduceți numărul contorului.", parent=win)
                return
            if not an_str or not luna_n:
                messagebox.showwarning("Selecție lipsă",
                    "Selectați anul și luna.", parent=win)
                return
            if not mzl_str:
                messagebox.showwarning("Câmp lipsă",
                    "Introduceți noua valoare MZL.", parent=win)
                return
            # Operatorul e preluat automat din Windows — nu poate fi gol

            try:
                mzl_nou = float(mzl_str.replace(",", ".").replace(" ", ""))
                if mzl_nou < 0:
                    raise ValueError("Negativ")
            except ValueError:
                messagebox.showerror("Valoare invalidă",
                    "MZL trebuie să fie un număr pozitiv.", parent=win)
                return

            try:
                an   = int(an_str)
                luna = LUNI_INV[luna_n]
            except (ValueError, KeyError):
                messagebox.showerror("Eroare", "An sau lună invalidă.", parent=win)
                return

            # Valoarea veche (manual sau calculată)
            manuale_existente = tdb.get_mzl_manual(ct)
            val_veche_man = manuale_existente.get((an, luna))

            if val_veche_man is not None:
                val_veche_str = f"{int(val_veche_man):,} (valoare manuală anterioară)"
            else:
                # MZL calculat automat — media zilnică (identic cu excel_report.py)
                try:
                    import calendar as _cal
                    rows_zi_s = tdb._conn().execute("""
                        SELECT zi, SUM(total_general) AS total_zi, COUNT(*) AS ore_zi
                        FROM inregistrari_orare
                        WHERE contor=? AND an=? AND luna=?
                        GROUP BY zi
                    """, (ct, an, luna)).fetchall()
                    zile_val_s = [r for r in rows_zi_s if int(r["ore_zi"] or 0) >= 22]
                    if zile_val_s:
                        _tot_s = sum(int(r["total_zi"] or 0) for r in zile_val_s)
                        _mzl_s = round(_tot_s / len(zile_val_s))
                        val_veche_str = (f"{_mzl_s:,} "
                                        f"({len(zile_val_s)} zile valide, calculat automat)")
                    else:
                        val_veche_str = "fără date valide"
                except Exception:
                    val_veche_str = "necunoscută"

            # ── Confirmare cu valoarea veche → nouă ──────────────────────────
            msg = (
                f"Contor:   {ct}\n"
                f"Perioadă: {luna_n} {an}\n"
                f"Localitate: {var_localitate.get()}\n\n"
                f"Valoare curentă:  {val_veche_str}\n"
                f"Valoare nouă:     {int(mzl_nou):,}\n\n"
                f"Operator: {op}\n\n"
                f"Confirmi modificarea?"
            )
            if not messagebox.askyesno("Confirmare modificare MZL", msg,
                                        parent=win):
                return

            # ── Verificare lock înainte de salvare ───────────────────────────
            tip_lock = f"Salvare MZL manual [{ct}] {luna_n} {an}"
            if not self._check_lock_and_warn(tip_lock):
                return

            # ── Salvare în SQLite — în thread separat ca să nu blocheze GUI ────
            btn_salveaza_ref = [None]  # referință mutabilă, sigură între thread-uri

            def _do_salvare():
                try:
                    tdb.upsert_mzl_manual(
                        contor=ct, an=an, luna=luna,
                        valoare=mzl_nou,
                        observatii=obs,
                        utilizator=op,
                    )
                    def _ok():
                        self._log(
                            f"  ✏️  MZL manual [{ct}] {luna_n} {an}: "
                            f"{val_veche_str} → {int(mzl_nou):,}  (operator: {op})"
                        )
                        try:
                            if win.winfo_exists():
                                messagebox.showinfo(
                                    "Salvat",
                                    f"MZL pentru contorul {ct} — {luna_n} {an}\n"
                                    f"a fost actualizat la {int(mzl_nou):,}.\n\n"
                                    f"Modificarea va fi aplicată la următoarea generare a raportului.",
                                    parent=win,
                                )
                                _on_close()
                        except Exception:
                            pass
                    win.after(0, _ok)
                except Exception as e:
                    def _err(err=str(e)):
                        try:
                            if btn_salveaza_ref[0] and btn_salveaza_ref[0].winfo_exists():
                                btn_salveaza_ref[0].configure(state="normal")
                        except Exception:
                            pass
                        try:
                            if win.winfo_exists():
                                messagebox.showerror("Eroare salvare",
                                    f"Nu s-a putut salva în baza de date:\n{err}", parent=win)
                        except Exception:
                            pass
                    win.after(0, _err)
                finally:
                    self._release_lock()  # eliberăm imediat după salvare

            threading.Thread(target=_do_salvare, daemon=True).start()

        _btn_s = _ctk_btn(btn_row, "💾  Salvează", on_salveaza, "success", width=150)
        _btn_s.pack(side="left", padx=8)
        # Setăm referința după creare, folosită în thread-ul de salvare
        try:
            btn_salveaza_ref  # verifică dacă variabila există în scope
            btn_salveaza_ref[0] = _btn_s
        except NameError:
            pass
        _ctk_btn(btn_row, "✖  Închide",   _on_close,  "secondary", width=130).pack(side="left", padx=8)

        # Focus pe câmpul contor la deschidere
        entry_contor.focus_set()
        win.bind("<Return>", lambda e: on_cauta_contor())


# ==============================================================================
# ENTRY POINT
# ==============================================================================
if __name__ == "__main__":
    if not CTK_AVAILABLE:
        print("[WARN] customtkinter nu este instalat.")
        print("       Ruleaza:  pip install customtkinter")
        print("       Aplicatia porneste cu interfata de rezerva (tkinter standard).\n")
    app = PeekApp()
    app.mainloop()
