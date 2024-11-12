# main.py
# Version v1.1
# Hlavní skript pro spuštění aplikace
# Alpha verze určená pro vytváření důladné analýzy reportů z excelu.

import ttkbootstrap as ttk
from gui import ExcelAnalyzerApp

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")  # Nastavíme téma aplikace
    app = ExcelAnalyzerApp(root)
    root.mainloop()
