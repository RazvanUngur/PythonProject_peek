# =============================================================================
# app.py — Fereastra principală PeekApp
# =============================================================================

import os
import re
import sys
import threading
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
    CENTRAL_FILE_FOLDER, CENTRAL_FILE_NAME, FONT_SMALL, FONT_MONO,
    BG_APP, BG_CARD, BG_LOG, BG_TITLE, FG_TITLE, RADIUS,
)
from bin_parser import quick_scan_bin, process_multiple_files
from log_parser import process_log_files
from centralizator import update_centralizator
from contoare_db import _load_contoare_db, _save_contoare_db, _delete_contor_from_centralizator
from gui_widgets import (
    _rr, _ctk_btn, _make_button, _ctk_frame, _ctk_label,
    _ctk_entry, _add_logo_header, _open_contor_dialog,
)


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

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _update_progress(self, percent, text):
        if CTK_AVAILABLE:
            self.progress.set(percent / 100)
        else:
            self.progress["value"] = percent
        self.lbl_status.configure(text=text)
        self.update_idletasks()

    def _log(self, msg):
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
    def _run_all_processing(self):
        has_bin = bool(self.selected_files)
        has_log = bool(self.selected_log_files)
        if not has_bin and not has_log:
            messagebox.showwarning("Atentie", "Selecteaza cel putin un fisier .bin sau .log!")
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
            for idx_r, r in enumerate(rezultate):
                if cancelled(): abort(); return          # ← între contoare centralizator
                pct_c = int(75 + (idx_r / n_contoare) * 20)
                self._update_progress(pct_c, f"Centralizator {r['id']} ({idx_r+1}/{n_contoare})…")
                try:
                    update_centralizator(r['path'], r['id'], CENTRAL_FILE_FOLDER)
                except Exception as e_c:
                    self._log(f"  ⚠ Centralizator eroare [{r['id']}]: {e_c}")

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
            out_dir = os.path.dirname(os.path.abspath(self.selected_files[0]))
            messagebox.showinfo(
                "Succes",
                f"Procesare finalizată!\n\n"
                f"Contoar{'e' if n_contoare!=1 else ''} procesat{'e' if n_contoare!=1 else ''}:   {n_contoare}\n"
                f"Fișiere:              {fisiere_valide}/{total}{goale_str}\n"
                f"Înregistrări orare:   {total_ore:,}\n\n"
                f"Rapoartele au fost salvate în:\n{out_dir}",
            )
        except Exception as e:
            import traceback
            self._log(f"\n✗ EROARE: {e}")
            traceback.print_exc()
            self._update_progress(0, "✗  Eroare la procesare.")
        finally:
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
                    for idx_r, r in enumerate(rezultate_bin):
                        if cancelled(): abort(); return
                        pct_c = int(bin_pct_end * 0.7 + (idx_r / n_c) * bin_pct_end * 0.28)
                        self._update_progress(pct_c,
                            f"Centralizator .bin {r['id']} ({idx_r+1}/{n_c})...")
                        try:
                            update_centralizator(r["path"], r["id"], CENTRAL_FILE_FOLDER)
                        except Exception as e_c:
                            self._log(f"  ⚠ Centralizator eroare [{r['id']}]: {e_c}")

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

                out_dir = os.path.dirname(
                    os.path.abspath(self.selected_log_files[0]))
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

                    for idx_r, r in enumerate(rezultate_log):
                        if cancelled(): abort(); return
                        pct_c = int(log_pct_start + 20 + (idx_r / n_cl) * 25)
                        self._update_progress(pct_c,
                            f"Centralizator .log {r['id']} ({idx_r+1}/{n_cl})...")
                        self._log(f"  📥 Centralizator ← [{r['id']}]  "
                                  f"{r['randuri']} ore  │  "
                                  f"B1={r['b1']:,}  B2={r['b2']:,}")
                        try:
                            update_centralizator(r["path"], r["id"], CENTRAL_FILE_FOLDER)
                        except Exception as e_c:
                            self._log(f"  ⚠ Centralizator eroare [{r['id']}]: {e_c}")

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

            out_path = (os.path.dirname(os.path.abspath(self.selected_files[0]))
                        if self.selected_files else
                        os.path.dirname(os.path.abspath(self.selected_log_files[0])))

            lines = ["Procesare finalizata!\n"]
            if rezultate_bin:
                lines.append(f"📦 .BIN — {len(rezultate_bin)} contoar{'e' if n_total!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_bin):,} ore)")
            if rezultate_log:
                lines.append(f"📦 .LOG — {len(rezultate_log)} contoar{'e' if n_total!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_log):,} ore)")
            lines.append(f"\nFisier{'e' if n_total!=1 else ''} salvat{'e' if n_total!=1 else ''} in:\n{out_path}")
            messagebox.showinfo("Succes", "\n".join(lines))

        except Exception as ex:
            self._log(f"\n✗ Eroare neasteptata: {ex}")
            self._update_progress(0, "✗  Eroare procesare.")
        finally:
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

                rezultate_bin = process_multiple_files(
                    self.selected_files,
                    stop_event=self.stop_event,
                    progress_callback=_bin_progress) or []
                if self.stop_event.is_set(): abort(); return

                if not rezultate_bin:
                    self._log("\n✗ Nicio data valida gasita in fisierele .bin.")
                else:
                    n_c = len(rezultate_bin)
                    self._update_progress(int(bin_pct_end * 0.65),
                        f"Actualizez centralizatorul ({n_c} contoar{'e' if n_c!=1 else ''} .bin)...")
                    for idx_r, r in enumerate(rezultate_bin):
                        if cancelled(): abort(); return
                        pct_c = int(bin_pct_end * 0.65 + (idx_r / n_c) * bin_pct_end * 0.30)
                        self._update_progress(pct_c,
                            f"Centralizator .bin {r['id']} ({idx_r+1}/{n_c})...")
                        try:
                            update_centralizator(r["path"], r["id"], CENTRAL_FILE_FOLDER)
                        except Exception as e_c:
                            self._log(f"  ⚠ Centralizator eroare [{r['id']}]: {e_c}")

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

                out_dir = os.path.dirname(os.path.abspath(self.selected_log_files[0]))

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

                rezultate_log = process_log_files(
                    self.selected_log_files,
                    output_dir=out_dir,
                    stop_event=self.stop_event,
                    progress_callback=_log_progress) or []

                if self.stop_event.is_set(): abort(); return

                if not rezultate_log:
                    self._log("\n✗ Nicio data valida gasita in fisierele .log.")
                else:
                    n_cl = len(rezultate_log)
                    self._update_progress(log_pct_start + 20,
                        f"Actualizez centralizatorul ({n_cl} contoar{'e' if n_cl!=1 else ''} .log)...")
                    self._log(f"\n  📊 Rapoarte .log generate: {n_cl} contoar{'e' if n_cl!=1 else ''}")

                    for idx_r, r in enumerate(rezultate_log):
                        if cancelled(): abort(); return
                        pct_c = int(log_pct_start + 20 + (idx_r / n_cl) * 25)
                        self._update_progress(pct_c,
                            f"Centralizator .log {r['id']} ({idx_r+1}/{n_cl})...")
                        self._log(f"  📥 Centralizator ← [{r['id']}]  "
                                  f"{r['randuri']} ore  │  "
                                  f"B1={r['b1']:,}  B2={r['b2']:,}")
                        try:
                            update_centralizator(r["path"], r["id"], CENTRAL_FILE_FOLDER)
                        except Exception as e_c:
                            self._log(f"  ⚠ Centralizator eroare [{r['id']}]: {e_c}")

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

            self._update_progress(100,
                f"✔  Finalizat — {n_total} contoar{'e' if n_total!=1 else ''} procesat{'e' if n_total!=1 else ''}.")

            out_path = (os.path.dirname(os.path.abspath(self.selected_files[0]))
                        if self.selected_files else
                        os.path.dirname(os.path.abspath(self.selected_log_files[0])))

            lines = ["Procesare finalizata!\n"]
            if rezultate_bin:
                lines.append(f"📦 .BIN — {len(rezultate_bin)} contoar{'e' if len(rezultate_bin)!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_bin):,} ore)")
            if rezultate_log:
                lines.append(f"📦 .LOG — {len(rezultate_log)} contoar{'e' if len(rezultate_log)!=1 else ''}  "
                             f"({sum(r['randuri'] for r in rezultate_log):,} ore)")
            lines.append(f"\nFisier{'e' if n_total!=1 else ''} salvat{'e' if n_total!=1 else ''} in:\n{out_path}")
            messagebox.showinfo("Succes", "\n".join(lines))

        except Exception as ex:
            self._log(f"\n✗ Eroare neasteptata: {ex}")
            self._update_progress(0, "✗  Eroare procesare.")
        finally:
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
        win.minsize(720, 460)
        W, H = 880, 580
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

        cols = ("Contor", "Drum", "Poziție km", "Localitate", "Tip", "IP")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                            height=14, style="Peek.Treeview")
        cw = {"Contor": 75, "Drum": 95, "Poziție km": 85,
              "Localitate": 140, "Tip": 180, "IP": 125}
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=cw.get(c, 100), anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
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

        def refresh_tree():
            for row in tree.get_children():
                tree.delete(row)
            db = _load_contoare_db()
            for ct, d in sorted(db.items()):
                tree.insert("", "end", values=(
                    ct, d.get("Drum",""), d.get("Pozitie_km",""),
                    d.get("Localitate",""), d.get("Tip",""), d.get("IP",""),
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
            db = _load_contoare_db()
            _open_contor_dialog(win, mode="edit", ct_id=ct_id,
                                data=db.get(ct_id,{}), on_save=refresh_tree,
                                app_ref=self)

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

        # ── Creare fereastră ──────────────────────────────────────────────────
        if CTK_AVAILABLE:
            win = ctk.CTkToplevel(self)
        else:
            win = tk.Toplevel(self)

        win.title("Procesare manuală MZL")
        win.resizable(False, False)
        win.grab_set()
        win.focus_force()

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
            # Localitate din ContoareDB
            loc = cdb.get_localitate(ct) if ct else ""
            var_localitate.set(loc if loc else "Necunoscută")

            # Ani disponibili din TrafficDB
            ani = tdb.get_ani_disponibili(ct)
            if not ani:
                messagebox.showwarning("Contor negăsit",
                    f"Contorul {ct} nu are date în baza de date.", parent=win)
                var_localitate.set("—")
                return

            combo_an.configure(values=[str(a) for a in ani])
            var_an.set(str(ani[-1]))   # selectăm cel mai recent an
            on_an_changed()            # populăm și lunile

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
            """Afișează MZL calculat și cel manual existent pentru an+lună."""
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

            # MZL calculat automat — media ZILNICĂ (suma totală / zile valide)
            # identic cu logica din excel_report.py (MIN_ORE_ZI ore/zi valide)
            import sqlite3, calendar
            try:
                conn = tdb._conn()
                # Pas 1: zile valide (≥22 ore înregistrate)
                rows_zi = conn.execute("""
                    SELECT zi, SUM(total_general) AS total_zi, COUNT(*) AS ore_zi
                    FROM inregistrari_orare
                    WHERE contor=? AND an=? AND luna=?
                    GROUP BY zi
                """, (ct, an, luna)).fetchall()

                zile_luna = calendar.monthrange(an, luna)[1]
                zile_valide = [r for r in rows_zi if int(r["ore_zi"] or 0) >= 22]
                n_zile_val  = len(zile_valide)

                if n_zile_val > 0:
                    total_zile_valide = sum(int(r["total_zi"] or 0) for r in zile_valide)
                    mzl_auto = round(total_zile_valide / n_zile_val)
                    var_mzl_calc.set(
                        f"{mzl_auto:,}  "
                        f"({n_zile_val}/{zile_luna} zile valide)")
                else:
                    var_mzl_calc.set("Fără date valide (< 22 ore/zi)")
            except Exception as e:
                var_mzl_calc.set(f"Eroare: {e}")

            # MZL manual existent
            manuale = tdb.get_mzl_manual(ct)
            val_man = manuale.get((an, luna))
            if val_man is not None:
                var_mzl_manual.set(f"{int(val_man):,}")
                var_mzl_nou.set(str(int(val_man)))
            else:
                var_mzl_manual.set("(nesetat)")
                var_mzl_nou.set("")

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

            # ── Salvare în SQLite ─────────────────────────────────────────────
            try:
                tdb.upsert_mzl_manual(
                    contor=ct, an=an, luna=luna,
                    valoare=mzl_nou,
                    observatii=obs,
                    utilizator=op,
                )
                self._log(
                    f"  ✏️  MZL manual [{ct}] {luna_n} {an}: "
                    f"{val_veche_str} → {int(mzl_nou):,}  (operator: {op})"
                )
                messagebox.showinfo(
                    "Salvat",
                    f"MZL pentru contorul {ct} — {luna_n} {an}\n"
                    f"a fost actualizat la {int(mzl_nou):,}.\n\n"
                    f"Modificarea va fi aplicată la următoarea generare a raportului.",
                    parent=win,
                )
                # Actualizăm afișarea valorii manuale în fereastră
                var_mzl_manual.set(f"{int(mzl_nou):,}")
                var_mzl_nou.set("")
            except Exception as e:
                messagebox.showerror("Eroare salvare",
                    f"Nu s-a putut salva în baza de date:\n{e}", parent=win)

        _ctk_btn(btn_row, "💾  Salvează",  on_salveaza,  "success",   width=150).pack(side="left", padx=8)
        _ctk_btn(btn_row, "✖  Închide",   win.destroy,  "secondary", width=130).pack(side="left", padx=8)

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
