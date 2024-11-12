# data_processing.py
# Modul pro zpracování dat

import pandas as pd
from tkinter import filedialog, messagebox

class DataProcessor:
    def __init__(self):
        self.data = None
        self.filtered_data = None

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                self.data = pd.read_excel(file_path, sheet_name='UN-023_QM_Auswertungen_GH', skiprows=1)
                self.data.columns = [
                    'Index', 'Datum Von', 'Linie', 'PPlatz', 'Storort',
                    'Storort Bezeichnung', 'Fab Nr.', 'Material Nr. Geraet',
                    'Geraet Bezeichnung', 'Material Nr.', 'Material Bezeichnung',
                    'Fehler', 'Fehler Bezeichnung', 'Kommentar'
                ]
                messagebox.showinfo("Úspěch", "Soubor byl úspěšně načten!")
            except Exception as e:
                messagebox.showerror("Chyba", f"Soubor se nepodařilo načíst: {e}")

    def apply_filters(self, filters):
        if self.data is not None:
            self.filtered_data = self.data
            for column, value in filters:
                if column and value:
                    self.filtered_data = self.filtered_data[self.filtered_data[column].astype(str) == value]
            messagebox.showinfo("Filtr", "Filtry byly aplikovány")
        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte soubor.")

    def get_columns(self):
        if self.data is not None:
            return list(self.data.columns)
        else:
            return []

    def get_filtered_data(self):
        return self.filtered_data if self.filtered_data is not None else self.data
