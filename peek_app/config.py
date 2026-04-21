# =============================================================================
# config.py — Constante și configurare globală
# =============================================================================

# ── Centralizator ─────────────────────────────────────────────────────────────
CENTRAL_FILE_NAME   = "0_Centralizator_PEEK-VEK.xlsx"
CENTRAL_FILE_FOLDER = r"L:\BIDMRCT\datePEEK"

# ── Mapare categorii vehicule → clase PEEK ────────────────────────────────────
VEHICLE_ANALYSIS = {
    "Autoturisme": ["Clasa_2"],
    "LGV":         ["Clasa_3"],
    "HGV":         ["Clasa_4", "Clasa_5", "Clasa_6", "Clasa_7"],
    "Autobuze":    ["Clasa_8"],
}

# ── Constante calcul ──────────────────────────────────────────────────────────
MIN_LUNI_AN      = 10  # luni valide minime pentru MZA completă
MIN_LUNI_AN_MAI  =  5  # luna Mai (nr. 5)  — fallback primar când < 10 luni valide
MIN_LUNI_AN_OCT  = 10  # luna Octombrie (nr. 10) — fallback secundar dacă Mai lipsește
MIN_ORE_ZI       = 22  # ore minime pentru zi validă
MIN_ZILE_LUNA    = 15  # zile valide minime pentru lună validă în MZA
MIN_ZILE_SAPT    =  7  # zile consecutive minime (alternativă la MIN_ZILE_LUNA)

# ── Culori grafice pentru benzi ───────────────────────────────────────────────
BAND_COLORS = ["2E75B6", "ED7D31", "A9D18E", "FF0000", "7030A0", "00B0F0"]

# ── GUI — teme și culori ──────────────────────────────────────────────────────
CTK_THEME = "system"   # "light" | "dark" | "system"
CTK_COLOR = "blue"     # "blue" | "green" | "dark-blue"

BG_APP   = "#F0F4F8"
BG_CARD  = "#FFFFFF"
BG_TITLE = "#1A2533"
FG_TITLE = "#FFFFFF"
BG_LOG   = "#F7F9FC"
RADIUS   = 12

# ── GUI — fonturi ─────────────────────────────────────────────────────────────
FONT_UI    = "Inter"
FONT_BTN   = ("Segoe UI", 13, "bold")
FONT_LBL   = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 10)
FONT_MONO  = ("Consolas", 10)
FONT_TITLE = ("Segoe UI", 14, "bold")

BTN_H      = 38
BTN_CORNER = 8

# ── GUI — culori butoane ──────────────────────────────────────────────────────
_BTN_COLORS = {
    "success":   ("#1A7A3C", "#25A853"),   # (normal, hover)
    "danger":    ("#C0392B", "#E74C3C"),
    "primary":   ("#1A5276", "#2471A3"),
    "secondary": ("#5D6D7E", "#717D8A"),
    "navy":      ("#0F3460", "#1A5276"),
    "info":      ("#117A65", "#17A589"),
}

# ── Baza de date contoare ─────────────────────────────────────────────────────
CONTOARE_COLS_ORDER = ["Drum", "Pozitie_km", "Localitate", "Tip", "IP"]
CONTOARE_HEADERS    = ["Contor", "Drum", "Poziție km", "Localitate", "Tip", "IP"]
TIP_OPTIONS = [
    "ADR 2000 - Clasificator",
    "ADR 2000 - WIM",
    "ADR 3000 - Clasificator",
    "ADR 3000 - WIM",
    "ADR Sabre - Clasificator",
    "ADR Sabre - WIM",
    "VEKs4",
]

# ── Mapare clase VEK ──────────────────────────────────────────────────────────
LOG_CLASS_MAP = {
    "Motorbike":          1,
    "Car":                2,
    "Car with trailer":   3,
    "Van":                4,
    "Lorry":              5,
    "Lorry with trailer": 6,
    "Truck":              7,
    "Bus":                8,
    "Other":             15,
}
