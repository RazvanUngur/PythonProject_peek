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

    def _try_round(self):
        try:
            import ctypes
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                self.winfo_id(), 33, ctypes.byref(ctypes.c_int(2)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

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

        # ── Gestionare Contoare ───────────────────────────────────────────────
        self.btn_contoare = _ctk_btn(self, "🗄️  Gestionare Contoare",
                                     self._open_contoare_manager, "navy", width=240)
        self.btn_contoare.pack(pady=(4, 6))

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

                out_dir = os.path.dirname(os.path.abspath(self.selected_log_files[0]))
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
