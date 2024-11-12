# dashboard.py
# Modul pro dashboard a analytiku

from tkinter import messagebox
from tkinter.ttk import Treeview, LabelFrame
from scipy.stats import pearsonr
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

class Dashboard:
    def __init__(self):
        self.stats_tree = None
        self.correlation_tree = None
        self.frequency_tree = None
        self.mini_graph_frame = None

    def setup_dashboard(self, stats_tree, correlation_tree, frequency_tree, mini_graph_frame):
        self.stats_tree = stats_tree
        self.correlation_tree = correlation_tree
        self.frequency_tree = frequency_tree
        self.mini_graph_frame = mini_graph_frame

    def update_dashboard(self, data, selected_column):
        if data is not None:
            # Vymažeme předchozí obsah
            self.stats_tree.delete(*self.stats_tree.get_children())
            self.correlation_tree.delete(*self.correlation_tree.get_children())
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

            numeric_columns = data.select_dtypes(include="number").columns

            # Základní statistiky
            for col in numeric_columns:
                mean = data[col].mean()
                median = data[col].median()
                std_dev = data[col].std()
                min_val = data[col].min()
                max_val = data[col].max()

                self.stats_tree.insert('', 'end', values=(f"{col} - Průměr", round(mean, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Medián", round(median, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Směrodatná odchylka", round(std_dev, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Minimum", round(min_val, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Maximum", round(max_val, 2)))

            # Korelační analýza
            for i in range(len(numeric_columns)):
                for j in range(i+1, len(numeric_columns)):
                    col1 = numeric_columns[i]
                    col2 = numeric_columns[j]
                    try:
                        corr_coef, _ = pearsonr(data[col1].dropna(), data[col2].dropna())
                        self.correlation_tree.insert('', 'end', values=(f"{col1} & {col2}", round(corr_coef, 2)))
                    except:
                        continue

            # Aktualizace četnostní tabulky
            if selected_column:
                self.update_frequency_table(selected_column, data)

            # Přidání dynamických grafů
            self.add_dynamic_graphs(data)

        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte data.")

    def update_frequency_table(self, selected_column, data):
        if selected_column and selected_column in data.columns:
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

            freq_series = data[selected_column].value_counts()

            for value, count in freq_series.items():
                self.frequency_tree.insert('', 'end', values=(str(value), count))

            # Přidání grafu četností do dynamických grafů
            fig, ax = plt.subplots(figsize=(5, 3))
            freq_series.plot(kind='bar', ax=ax)
            ax.set_title(f"Četnost hodnot ve sloupci {selected_column}")
            ax.set_xlabel("Hodnota")
            ax.set_ylabel("Počet výskytů")
            plt.xticks(rotation=45)
            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=self.mini_graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
        else:
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

    def add_dynamic_graphs(self, data):
        # Vymažeme předchozí grafy
        for widget in self.mini_graph_frame.winfo_children():
            widget.destroy()

        # Přidáme další dynamické grafy
        if 'Datum Von' in data.columns and 'Fehler' in data.columns:
            fig2, ax2 = plt.subplots(figsize=(5, 3))
            # Convert dates to datetime format
            data['Datum Von'] = pd.to_datetime(data['Datum Von'], errors='coerce')
            # Group by date and count occurrences of "Fehler" per date
            data.set_index('Datum Von').resample('D')['Fehler'].count().plot(ax=ax2)
            ax2.set_title("Frekvence chyb v čase")
            ax2.set_xlabel("Datum")
            ax2.set_ylabel("Počet chyb")
            plt.tight_layout()
            canvas2 = FigureCanvasTkAgg(fig2, master=self.mini_graph_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill='both', expand=True)

        if 'PPlatz' in data.columns:
            fig3, ax3 = plt.subplots(figsize=(5, 3))
            # Count occurrences of each unique value in "PPlatz"
            data['PPlatz'].value_counts().plot(kind='bar', ax=ax3)
            ax3.set_title("Četnost hodnot v PPlatz")
            ax3.set_xlabel("PPlatz")
            ax3.set_ylabel("Četnost")
            plt.tight_layout()
            canvas3 = FigureCanvasTkAgg(fig3, master=self.mini_graph_frame)
            canvas3.draw()
            canvas3.get_tk_widget().pack(fill='both', expand=True)
