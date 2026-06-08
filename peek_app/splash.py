# splash.py — ecran de încărcare cu colțuri rotunjite și temă system
import tkinter as tk
import sys
import os


def _get_system_theme():
    """Detectează light/dark din registry Windows."""
    try:
        import winreg
        reg = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        val, _ = winreg.QueryValueEx(reg, "AppsUseLightTheme")
        winreg.CloseKey(reg)
        return "light" if val == 1 else "dark"
    except Exception:
        return "light"


def _rounded_rect(canvas, x1, y1, x2, y2, r, **kwargs):
    """Desenează un dreptunghi cu colțuri rotunjite pe canvas."""
    points = [
        x1+r, y1,   x2-r, y1,
        x2,   y1,   x2,   y1+r,
        x2,   y2-r, x2,   y2,
        x2-r, y2,   x1+r, y2,
        x1,   y2,   x1,   y2-r,
        x1,   y1+r, x1,   y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


class SplashScreen:

    # Palette light / dark
    THEMES = {
        "dark": {
            "bg_outer":  "#0F172A",   # fundal fereastra (în spatele canvas)
            "bg_card":   "#1E293B",   # cardul rotunjit
            "fg_title":  "#F1F5F9",
            "fg_sub":    "#94A3B8",
            "bar_bg":    "#334155",
            "bar_fill":  "#3B82F6",
            "border":    "#334155",
        },
        "light": {
            "bg_outer":  "#E2E8F0",
            "bg_card":   "#FFFFFF",
            "fg_title":  "#1E293B",
            "fg_sub":    "#64748B",
            "bar_bg":    "#CBD5E1",
            "bar_fill":  "#2563EB",
            "border":    "#CBD5E1",
        },
    }

    def __init__(self):
        theme = _get_system_theme()
        self.c = self.THEMES[theme]

        W, H = 380, 180
        R = 20          # raza colțurilor cardului

        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg=self.c["bg_outer"])

        # Transparență pe Windows — culoarea bg_outer devine transparentă
        # astfel cardul rotunjit pare că plutește
        try:
            self.root.wm_attributes("-transparentcolor", self.c["bg_outer"])
        except Exception:
            pass

        # Centrat pe ecran
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")

        # Canvas principal
        self.cv = tk.Canvas(self.root, width=W, height=H,
                            bg=self.c["bg_outer"],
                            highlightthickness=0)
        self.cv.pack(fill="both", expand=True)

        # Card rotunjit
        _rounded_rect(self.cv, 10, 10, W-10, H-10, R,
                      fill=self.c["bg_card"],
                      outline=self.c["border"],
                      width=1)

        # Titlu
        self.cv.create_text(W//2, 52,
                            text="PEEK Traffic Analyzer",
                            font=("Segoe UI", 17, "bold"),
                            fill=self.c["fg_title"],
                            anchor="center")

        # Status
        self.txt_status = self.cv.create_text(W//2, 84,
                            text="Se inițializează...",
                            font=("Segoe UI", 10),
                            fill=self.c["fg_sub"],
                            anchor="center")

        # Bara de progres — fundal
        BAR_X1, BAR_Y = 40, 118
        BAR_W, BAR_H   = W - 80, 7
        BAR_R = BAR_H // 2
        _rounded_rect(self.cv, BAR_X1, BAR_Y,
                      BAR_X1 + BAR_W, BAR_Y + BAR_H, BAR_R,
                      fill=self.c["bar_bg"], outline="")

        # Bara de progres — fill (începe la 0)
        self._bar_x1 = BAR_X1
        self._bar_y  = BAR_Y
        self._bar_w  = BAR_W
        self._bar_h  = BAR_H
        self._bar_r  = BAR_R
        self._bar_fill_id = None
        self._draw_bar(0)

        # Versiune
        self.cv.create_text(W//2, H - 22,
                            text="v7.01",
                            font=("Segoe UI", 8),
                            fill=self.c["fg_sub"],
                            anchor="center")

        # Drag cu mouse (opțional — poate muta fereastra)
        self.cv.bind("<ButtonPress-1>",   self._drag_start)
        self.cv.bind("<B1-Motion>",       self._drag_move)
        self._drag_x = self._drag_y = 0

        self.root.update()

    def _draw_bar(self, pct):
        """Redesenează bara de progres la procentul dat (0-100)."""
        if self._bar_fill_id:
            self.cv.delete(self._bar_fill_id)
        if pct <= 0:
            return
        fill_w = max(self._bar_r * 2, int(self._bar_w * pct / 100))
        self._bar_fill_id = _rounded_rect(
            self.cv,
            self._bar_x1, self._bar_y,
            self._bar_x1 + fill_w, self._bar_y + self._bar_h,
            self._bar_r,
            fill=self.c["bar_fill"], outline="")

    def set_status(self, text, progress=None):
        self.cv.itemconfig(self.txt_status, text=text)
        if progress is not None:
            self._draw_bar(progress)
        self.root.update()

    def _drag_start(self, e):
        self._drag_x, self._drag_y = e.x, e.y

    def _drag_move(self, e):
        dx = e.x - self._drag_x
        dy = e.y - self._drag_y
        x = self.root.winfo_x() + dx
        y = self.root.winfo_y() + dy
        self.root.geometry(f"+{x}+{y}")

    def close(self):
        self.root.destroy()
