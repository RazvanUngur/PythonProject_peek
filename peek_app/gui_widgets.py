# =============================================================================
# gui_widgets.py — Widget-uri și helper-e GUI reutilizabile
# =============================================================================
# Exportă:
#   _rr(), _ctk_btn(), _make_button(), _ctk_frame(), _ctk_label(),
#   _ctk_entry(), _add_logo_header(), _open_contor_dialog()
# =============================================================================

import os
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox

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
    BG_APP, BG_CARD, BG_TITLE, FG_TITLE, BG_LOG, RADIUS,
    FONT_SMALL, FONT_MONO, FONT_UI, FONT_BTN, FONT_LBL, FONT_TITLE,
    BTN_H, BTN_CORNER, _BTN_COLORS,
    CTK_THEME, CTK_COLOR,
    CONTOARE_COLS_ORDER, CONTOARE_HEADERS, TIP_OPTIONS,
)
from contoare_db import _load_contoare_db, _save_contoare_db



def _rr(canvas, x1, y1, x2, y2, r=RADIUS, **kw):
    """Desenează un dreptunghi cu colțuri rotunjite pe canvas."""
    pts = [
        x1+r, y1,   x2-r, y1,
        x2,   y1,   x2,   y1+r,
        x2,   y2-r, x2,   y2,
        x2-r, y2,   x1+r, y2,
        x1,   y2,   x1,   y2-r,
        x1,   y1+r, x1,   y1,
        x1+r, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, **kw)



# ==============================================================================
# GUI — CustomTkinter (aspect Windows 11 modern, butoane rotunjite nativ)
# ==============================================================================
# Instalare (o singura data):  pip install customtkinter
# ==============================================================================

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except ImportError:
    CTK_AVAILABLE = False
    import tkinter.ttk as _ttk_fallback

# Tema si culori
CTK_THEME   = "system"      # "light" | "dark" | "system"
CTK_COLOR   = "blue"        # tema de accent: "blue" | "green" | "dark-blue"

# Paleta manuala pentru butoane colorate
_BTN_COLORS = {
    "success":   ("#1A7A3C", "#25A853"),   # (normal, hover)
    "danger":    ("#C0392B", "#E74C3C"),
    "primary":   ("#1A5276", "#2471A3"),
    "secondary": ("#5D6D7E", "#717D8A"),
    "navy":      ("#0F3460", "#1A5276"),
    "info":      ("#117A65", "#17A589"),
}

# Dimensiuni comune
BTN_H       = 38
BTN_CORNER  = 8
FONT_BTN    = ("Segoe UI", 13, "bold")
FONT_LBL    = ("Segoe UI", 12)
FONT_SMALL  = ("Segoe UI", 10)
FONT_MONO   = ("Consolas", 10)
FONT_TITLE  = ("Segoe UI", 14, "bold")


def _ctk_btn(parent, text, command, ck="primary", width=200, state="normal", **kw):
    """Buton CTk colorat cu colturi rotunjite. Functioneaza fara canvas tricks."""
    if not CTK_AVAILABLE:
        # fallback tk.Button daca customtkinter nu e instalat
        bg, hbg = _BTN_COLORS.get(ck, _BTN_COLORS["primary"])
        b = tk.Button(parent, text=text, command=command,
                      bg=bg, fg="white", activebackground=hbg,
                      font=FONT_BTN, relief="flat", cursor="hand2",
                      state=state)
        b.bind("<Enter>", lambda e: b.config(bg=hbg))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b
    fg, hfg = _BTN_COLORS.get(ck, _BTN_COLORS["primary"])
    return ctk.CTkButton(
        parent, text=text, command=command,
        fg_color=fg, hover_color=hfg, text_color="white",
        width=width, height=BTN_H,
        corner_radius=BTN_CORNER,
        font=FONT_BTN,
        state=state,
        **kw
    )


def _make_button(parent, text, command, style_ttk="", style_tk="",
                 width=None, state="normal", color_key=None):
    """Factory compatibil cu restul codului — mapeaza style_ttk pe CTkButton."""
    ck  = color_key or style_ttk or "primary"
    # width in caractere → pixeli aproximativi
    wpx = (width * 9 + 20) if width else max(len(text) * 9 + 30, 160)
    return _ctk_btn(parent, text=text, command=command,
                    ck=ck, width=wpx, state=state)


def _ctk_frame(parent, **kw):
    """Frame CTk cu colturi rotunjite si fundal card."""
    if not CTK_AVAILABLE:
        return tk.Frame(parent, bg="#FFFFFF", relief="solid", bd=1, **kw)
    return ctk.CTkFrame(parent, corner_radius=10, **kw)


def _ctk_label(parent, text, font=None, text_color=None, **kw):
    if not CTK_AVAILABLE:
        return tk.Label(parent, text=text, font=font or FONT_LBL,
                        fg=text_color or "#1A1A2E", bg="#F5F5F5", **kw)
    return ctk.CTkLabel(parent, text=text, font=font or FONT_LBL,
                        text_color=text_color or "#1A1A2E", **kw)


def _ctk_entry(parent, textvariable=None, width=200, state="normal", **kw):
    if not CTK_AVAILABLE:
        return ttk.Entry(parent, textvariable=textvariable, width=width//8, **kw)
    return ctk.CTkEntry(parent, textvariable=textvariable,
                        width=width, height=32, corner_radius=6,
                        state=state, **kw)


# Dialog adaugare/editare contor — rescris pentru CTk
def _add_logo_header(win, title_txt, app_ref=None):
    """
    Adauga un header cu logo mic stanga + titlu dreapta, identic in toate ferestrele.
    app_ref: referinta la PeekApp pentru a accesa imaginea logo deja incarcata.
    """
    if CTK_AVAILABLE:
        header = ctk.CTkFrame(win, fg_color="#FFFFFF", corner_radius=0, height=56)
    else:
        header = tk.Frame(win, bg="#FFFFFF", height=56)
    header.pack(fill="x", padx=0, pady=(0, 0))
    header.pack_propagate(False)

    # Logo mic stanga (40x40) — din imaginea deja incarcata in app
    logo_shown = False
    if app_ref is not None and hasattr(app_ref, '_logo') and app_ref._logo is not None:
        try:
            if CTK_AVAILABLE and hasattr(app_ref._logo, '_light_image'):
                # Cream o versiune mica din imaginea originala
                orig = app_ref._logo._light_image
                thumb = orig.copy()
                thumb.thumbnail((120, 40), Image.Resampling.LANCZOS)
                small_img = ctk.CTkImage(light_image=thumb, dark_image=thumb,
                                         size=(thumb.width, thumb.height))
                lbl = ctk.CTkLabel(header, image=small_img, text="",
                                   fg_color="transparent")
                lbl.pack(side="left", padx=(12, 8), pady=8)
                lbl._img_ref = small_img   # previne garbage collection
                logo_shown = True
            elif not CTK_AVAILABLE and isinstance(app_ref._logo, ImageTk.PhotoImage):
                tk.Label(header, image=app_ref._logo, bg="#FFFFFF").pack(
                    side="left", padx=(12, 8), pady=8)
                logo_shown = True
        except Exception as ex:
            print(f"[WARN] logo header: {ex}")

    # Titlu fereastra (langa logo sau singur daca logo lipseste)
    padx_title = (8, 16) if logo_shown else (16, 16)
    if CTK_AVAILABLE:
        ctk.CTkLabel(header, text=title_txt, font=FONT_TITLE,
                     fg_color="transparent",
                     text_color="#1A2533").pack(side="left", padx=padx_title, pady=8)
    else:
        tk.Label(header, text=title_txt, font=FONT_TITLE,
                 bg="#FFFFFF", fg="#1A2533").pack(side="left", padx=padx_title, pady=8)

    # Linie separatoare sub header
    if CTK_AVAILABLE:
        ctk.CTkFrame(win, height=1, fg_color="#CCCCCC").pack(fill="x", pady=(0, 6))
    else:
        ttk.Separator(win, orient="horizontal").pack(fill="x", pady=(0, 6))


def _open_contor_dialog(parent, mode="add", ct_id=None, data=None, on_save=None, app_ref=None):
    if data is None:
        data = {}

    if CTK_AVAILABLE:
        win = ctk.CTkToplevel(parent)
    else:
        win = tk.Toplevel(parent)

    title_txt = "Adaugă contoar" if mode == "add" else f"Editează contoar {ct_id}"
    win.title(title_txt)
    win.resizable(False, False)
    win.grab_set()
    win.focus_force()

    W, H = 460, 560
    sw = win.winfo_screenwidth(); sh = win.winfo_screenheight()
    win.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

    # Header cu logo + titlu
    _add_logo_header(win, title_txt, app_ref=app_ref)

    # Grid campuri
    grid = ctk.CTkFrame(win, fg_color="transparent") if CTK_AVAILABLE else tk.Frame(win)
    grid.pack(fill="x", padx=24, pady=4)

    vars_ = {}

    def add_field(label, key, row, initial="", readonly=False):
        if CTK_AVAILABLE:
            ctk.CTkLabel(grid, text=label, font=FONT_SMALL,
                         anchor="e", width=110).grid(row=row, column=0, padx=(0,8), pady=6, sticky="e")
        else:
            tk.Label(grid, text=label, font=FONT_SMALL,
                     anchor="e", width=14).grid(row=row, column=0, padx=(0,8), pady=6, sticky="e")
        var = tk.StringVar(value=initial)
        if CTK_AVAILABLE:
            e = ctk.CTkEntry(grid, textvariable=var, width=240, height=30,
                             corner_radius=6,
                             state="disabled" if readonly else "normal")
        else:
            e = ttk.Entry(grid, textvariable=var, width=28,
                          state="readonly" if readonly else "normal")
        e.grid(row=row, column=1, pady=6, sticky="w")
        vars_[key] = var

    add_field("Nr. contor *", "ct",  0, ct_id or "", readonly=(mode=="edit"))
    add_field("Drum",         "drum",1, data.get("Drum",""))
    add_field("Poziție km",   "poz", 2, data.get("Pozitie_km",""))
    add_field("Localitate",   "loc", 3, data.get("Localitate",""))

    # Tip dropdown
    if CTK_AVAILABLE:
        ctk.CTkLabel(grid, text="Tip", font=FONT_SMALL,
                     anchor="e", width=110).grid(row=4, column=0, padx=(0,8), pady=6, sticky="e")
    else:
        tk.Label(grid, text="Tip", font=FONT_SMALL,
                 anchor="e", width=14).grid(row=4, column=0, padx=(0,8), pady=6, sticky="e")

    # La adaugare: Tip gol; la editare: valoarea existenta din DB
    tip_initial = data.get("Tip", "") if mode == "edit" else ""
    var_tip = tk.StringVar(value=tip_initial)
    if CTK_AVAILABLE:
        ctk.CTkComboBox(grid, variable=var_tip, values=TIP_OPTIONS,
                        width=240, height=30, state="readonly",
                        corner_radius=6).grid(row=4, column=1, pady=6, sticky="w")
    else:
        ttk.Combobox(grid, textvariable=var_tip, values=TIP_OPTIONS,
                     state="readonly", width=28).grid(row=4, column=1, pady=6, sticky="w")
    vars_["tip"] = var_tip

    add_field("IP", "ip", 5, data.get("IP",""))

    # DRDP dropdown
    if CTK_AVAILABLE:
        ctk.CTkLabel(grid, text="DRDP", font=FONT_SMALL,
                     anchor="e", width=110).grid(row=6, column=0, padx=(0,8), pady=6, sticky="e")
    else:
        tk.Label(grid, text="DRDP", font=FONT_SMALL,
                 anchor="e", width=14).grid(row=6, column=0, padx=(0,8), pady=6, sticky="e")

    DRDP_OPTIONS = ["", "București", "Craiova", "Timișoara", "Cluj",
                    "Brașov", "Iași", "Constanța", "Buzău"]
    drdp_initial = data.get("DRDP", "") or data.get("drdp", "")
    var_drdp = tk.StringVar(value=drdp_initial)
    if CTK_AVAILABLE:
        ctk.CTkComboBox(grid, variable=var_drdp, values=DRDP_OPTIONS,
                        width=240, height=30, state="readonly",
                        corner_radius=6).grid(row=6, column=1, pady=6, sticky="w")
    else:
        ttk.Combobox(grid, textvariable=var_drdp, values=DRDP_OPTIONS,
                     state="readonly", width=28).grid(row=6, column=1, pady=6, sticky="w")
    vars_["drdp"] = var_drdp

    # Lat / Lng — câmpuri numerice libere
    def _coord_initial(key_upper, key_lower):
        v = data.get(key_upper) or data.get(key_lower) or ""
        try:
            return f"{float(v):.6f}" if v not in ("", None) else ""
        except Exception:
            return str(v) if v else ""

    add_field("Latitudine", "lat", 7, _coord_initial("lat", "Lat"))
    add_field("Longitudine", "lng", 8, _coord_initial("lng", "Lng"))

    # Separator + butoane
    if CTK_AVAILABLE:
        ctk.CTkFrame(win, height=1, fg_color="#CCCCCC").pack(fill="x", padx=20, pady=(10, 4))
    else:
        ttk.Separator(win, orient="horizontal").pack(fill="x", padx=20, pady=(10, 4))

    btn_row = ctk.CTkFrame(win, fg_color="transparent") if CTK_AVAILABLE else tk.Frame(win)
    btn_row.pack(pady=10)

    def on_save_click():
        ct_val = vars_["ct"].get().strip()
        if not ct_val:
            messagebox.showwarning("Câmp obligatoriu",
                                   "Numărul contoarului este obligatoriu!", parent=win)
            return
        new_data = {
            "Drum":       vars_["drum"].get().strip(),
            "Pozitie_km": vars_["poz"].get().strip(),
            "Localitate": vars_["loc"].get().strip(),
            "Tip":        vars_["tip"].get().strip(),
            "IP":         vars_["ip"].get().strip(),
            "DRDP":       vars_["drdp"].get().strip(),
        }
        # Validare și conversie coordonate
        for coord_key, var_key in [("lat", "lat"), ("lng", "lng")]:
            raw = vars_[var_key].get().strip().replace(",", ".")
            if raw:
                try:
                    new_data[coord_key] = float(raw)
                except ValueError:
                    messagebox.showerror(
                        "Coordonată invalidă",
                        f"Valoarea pentru {'Latitudine' if coord_key=='lat' else 'Longitudine'} "
                        f"nu este un număr valid: {raw!r}",
                        parent=win)
                    return
            else:
                new_data[coord_key] = None
        db = _load_contoare_db()
        if mode == "add" and ct_val in db:
            existing = db[ct_val]
            msg = (f"Contoarul {ct_val} există deja:\n"
                   f"  {existing.get('Drum','')} {existing.get('Pozitie_km','')} "
                   f"{existing.get('Localitate','')}\n\nVrei să editezi datele existente?")
            if messagebox.askyesno("Contoar existent", msg, parent=win):
                win.destroy()
                _open_contor_dialog(parent, mode="edit", ct_id=ct_val,
                                    data=existing, on_save=on_save, app_ref=app_ref)
            return
        db[ct_val] = new_data

        # ── Scriere în Excel (centralizator) ──────────────────────────────
        _save_contoare_db(db)

        # ── Scriere în SQLite contoare.db ─────────────────────────────────
        try:
            from database import get_contoare_db
            get_contoare_db().upsert(ct_val, new_data)
        except Exception as _e:
            print(f"[WARN] contoare.db write eroare [{ct_val}]: {_e}")

        action = "adăugat" if mode == "add" else "actualizat"
        messagebox.showinfo("Salvat", f"Contoarul {ct_val} a fost {action}.", parent=win)
        win.destroy()
        if on_save:
            on_save()

    _ctk_btn(btn_row, "💾  Salvează", on_save_click, "success", width=150).pack(side="left", padx=8)
    _ctk_btn(btn_row, "✖  Anulează", win.destroy,   "secondary", width=140).pack(side="left", padx=8)


# ==============================================================================
# PeekApp — fereastra principala CTk
# ==============================================================================