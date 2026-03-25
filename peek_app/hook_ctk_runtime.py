# hook_ctk_runtime.py
# Acest fișier e rulat de PyInstaller ÎNAINTE de orice alt cod
# Setează variabila de mediu ca customtkinter să găsească assets-urile

import os
import sys

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # Suntem în executabilul PyInstaller
    # Setăm calea către assets-urile customtkinter
    ctk_path = os.path.join(sys._MEIPASS, 'customtkinter')
    os.environ['CTK_ASSETS'] = ctk_path
    
    # Adăugăm _MEIPASS în sys.path ca import să funcționeze
    if sys._MEIPASS not in sys.path:
        sys.path.insert(0, sys._MEIPASS)
