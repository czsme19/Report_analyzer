import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Treeview, Scrollbar, LabelFrame
from scipy.stats import pearsonr
import datetime

class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Analyzer")
        self.data = None
        self.filtered_data = None

        # Nastavení tématu
        style = ttk.Style("flatly")

        # Vytvoření záložek
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=BOTH, expand=True)

        # Tab pro načítání souboru (vylepšená úvodní stránka)
        self.load_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.load_tab, text='Welcome')

        # Tab pro filtrování
        self.filter_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.filter_tab, text='Filters')

        # Tab pro grafy
        self.plot_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.plot_tab, text='Plots')

        # Tab pro dashboard
        self.dashboard_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.dashboard_tab, text='Dashboard')

        self.create_load_tab()
        self.create_filter_tab()
        self.create_plot_tab()
        self.create_dashboard_tab()

    def create_load_tab(self):
        # Vylepšená úvodní stránka
        self.welcome_frame = ttk.Frame(self.load_tab)
        self.welcome_frame.pack(fill=BOTH, expand=True, pady=50)

        # Dynamický uvítací text na základě denní doby
        current_hour = datetime.datetime.now().hour
        if current_hour < 12:
            greeting = "Dobré ráno"
        elif current_hour < 18:
            greeting = "Dobrý den"
        else:
            greeting = "Dobrý večer"

        self.title_label = ttk.Label(self.welcome_frame, text=f"{greeting}, vítejte v aplikaci Advanced Excel Analyzer", font=("Inter", 24, "bold"))
        self.title_label.pack(pady=20)

        # Popis aplikace
        self.description_label = ttk.Label(self.welcome_frame, text="Analyzujte svá Excel data snadno a rychle. Nahrajte soubor a začněte.", font=("Inter", 14))
        self.description_label.pack(pady=10)

        # Tip dne pro uživatele
        self.tip_label = ttk.Label(self.welcome_frame, text="Tip dne: Používejte filtry pro rychlé nalezení klíčových informací ve vašich datech.", font=("Inter", 12, "italic"), foreground="gray")
        self.tip_label.pack(pady=10)

        # Tlačítko pro načtení souboru
        self.file_button = ttk.Button(self.welcome_frame, text="Načíst Excel soubor", command=self.load_file, style="primary.TButton", bootstyle=PRIMARY)
        self.file_button.pack(pady=20)
        self.file_button_tooltip = ttk.Label(self.welcome_frame, text="Podporované formáty: .xlsx, .xls", font=("Inter", 10), foreground="gray")
        self.file_button_tooltip.pack(pady=5)

    def create_filter_tab(self):
        # Multi-level Filter options with dynamic value dropdown
        self.filter_frame = ttk.Frame(self.filter_tab)
        self.filter_frame.pack(pady=10, fill=X)

        self.filter_label = ttk.Label(self.filter_frame, text="Možnosti filtrování:", font=("Inter", 12, "bold"))
        self.filter_label.pack(anchor=W, padx=10)

        # Frame to hold dynamic filters and buttons
        self.filter_controls_frame = ttk.Frame(self.filter_tab)
        self.filter_controls_frame.pack(pady=5, fill=X)

        # Frame for buttons
        self.filter_buttons_frame = ttk.Frame(self.filter_controls_frame)
        self.filter_buttons_frame.pack(side=TOP, fill=X, padx=10, pady=5)

        self.add_filter_button = ttk.Button(self.filter_buttons_frame, text="+ Přidat filtr", command=self.add_filter, bootstyle=SUCCESS)
        self.add_filter_button.pack(side=LEFT, padx=5)

        self.remove_filter_button = ttk.Button(self.filter_buttons_frame, text="- Odebrat filtr", command=self.remove_filter, bootstyle=DANGER)
        self.remove_filter_button.pack(side=LEFT, padx=5)

        self.apply_filter_button = ttk.Button(self.filter_buttons_frame, text="Použít filtry", command=self.apply_filters, bootstyle=PRIMARY)
        self.apply_filter_button.pack(side=LEFT, padx=5)

        # Frame to hold dynamic filters
        self.dynamic_filters_frame = ttk.Frame(self.filter_tab)
        self.dynamic_filters_frame.pack(fill=X, padx=10, pady=5)

        self.filters = []
        self.add_filter()  # Přidáme první filtr

        # Table to display filtered rows
        self.table_frame = ttk.Frame(self.filter_tab)
        self.table_frame.pack(fill=BOTH, expand=True)
        self.tree = Treeview(self.table_frame)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=RIGHT, fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

    def add_filter(self):
        filter_row = ttk.Frame(self.dynamic_filters_frame)
        filter_row.pack(fill=X, pady=2)

        column_combobox = Combobox(filter_row, state="readonly")
        column_combobox.pack(side=LEFT, padx=5, pady=5)
        value_combobox = Combobox(filter_row, state="readonly")
        value_combobox.pack(side=LEFT, padx=5, pady=5)
        column_combobox.bind("<<ComboboxSelected>>", lambda e, vc=value_combobox, cc=column_combobox: self.update_value_options(vc, cc))

        self.filters.append((column_combobox, value_combobox))

        if self.data is not None:
            columns = list(self.data.columns)
            column_combobox['values'] = columns

    def remove_filter(self):
        if self.filters:
            filter_widgets = self.filters.pop()
            filter_widgets[0].master.destroy()  # Zničíme celý řádek s filtry

    def create_plot_tab(self):
        # Plot options with customizable X and Y axes
        self.plot_options_frame = ttk.Frame(self.plot_tab)
        self.plot_options_frame.pack(pady=10)

        self.plot_label = ttk.Label(self.plot_options_frame, text="Možnosti grafu:", font=("Inter", 12, "bold"))
        self.plot_label.grid(row=0, columnspan=2, sticky=W, padx=10)

        self.plot_x_label = ttk.Label(self.plot_options_frame, text="Vyberte sloupec pro osu X:")
        self.plot_x_label.grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.plot_x_combobox = Combobox(self.plot_options_frame, state="readonly", width=30)
        self.plot_x_combobox.grid(row=1, column=1, padx=5, pady=5)

        self.plot_y_label = ttk.Label(self.plot_options_frame, text="Vyberte sloupec pro osu Y:")
        self.plot_y_label.grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.plot_y_combobox = Combobox(self.plot_options_frame, state="readonly", width=30)
        self.plot_y_combobox.grid(row=2, column=1, padx=5, pady=5)

        # Výběr typu grafu
        self.plot_type_label = ttk.Label(self.plot_options_frame, text="Vyberte typ grafu:")
        self.plot_type_label.grid(row=3, column=0, sticky='e', padx=5, pady=5)
        self.plot_type_combobox = Combobox(self.plot_options_frame, state="readonly", width=30)
        self.plot_type_combobox['values'] = ['Scatter', 'Line', 'Bar', 'Histogram']
        self.plot_type_combobox.current(0)
        self.plot_type_combobox.grid(row=3, column=1, padx=5, pady=5)

        # Přidání dalších možností filtrování pro graf
        self.plot_filter_frame = ttk.Frame(self.plot_tab)
        self.plot_filter_frame.pack(pady=10, fill=X)

        self.plot_filter_label = ttk.Label(self.plot_filter_frame, text="Filtry pro graf:", font=("Inter", 12, "bold"))
        self.plot_filter_label.pack(anchor=W, padx=10)

        # Frame for plot filter buttons
        self.plot_filter_buttons_frame = ttk.Frame(self.plot_filter_frame)
        self.plot_filter_buttons_frame.pack(fill=X, padx=10, pady=5)

        self.add_plot_filter_button = ttk.Button(self.plot_filter_buttons_frame, text="+ Přidat filtr", command=self.add_plot_filter, bootstyle=SUCCESS)
        self.add_plot_filter_button.pack(side=LEFT, padx=5)

        self.remove_plot_filter_button = ttk.Button(self.plot_filter_buttons_frame, text="- Odebrat filtr", command=self.remove_plot_filter, bootstyle=DANGER)
        self.remove_plot_filter_button.pack(side=LEFT, padx=5)

        # Frame to hold dynamic plot filters
        self.dynamic_plot_filters_frame = ttk.Frame(self.plot_filter_frame)
        self.dynamic_plot_filters_frame.pack(fill=X, padx=10, pady=5)

        self.plot_filters = []
        self.add_plot_filter()  # Přidáme první plot filtr

        self.plot_button = ttk.Button(self.plot_tab, text="Generovat graf", command=self.plot_data, bootstyle=PRIMARY)
        self.plot_button.pack(pady=10)

    def add_plot_filter(self):
        plot_filter_row = ttk.Frame(self.dynamic_plot_filters_frame)
        plot_filter_row.pack(fill=X, pady=2)

        plot_column_combobox = Combobox(plot_filter_row, state="readonly")
        plot_column_combobox.pack(side=LEFT, padx=5, pady=5)
        plot_value_combobox = Combobox(plot_filter_row, state="readonly")
        plot_value_combobox.pack(side=LEFT, padx=5, pady=5)
        plot_column_combobox.bind("<<ComboboxSelected>>", lambda e, vc=plot_value_combobox, cc=plot_column_combobox: self.update_value_options(vc, cc))

        self.plot_filters.append((plot_column_combobox, plot_value_combobox))

        if self.data is not None:
            columns = list(self.data.columns)
            plot_column_combobox['values'] = columns

    def remove_plot_filter(self):
        if self.plot_filters:
            plot_filter_widgets = self.plot_filters.pop()
            plot_filter_widgets[0].master.destroy()  # Zničíme celý řádek s plot filtry

    def create_dashboard_tab(self):
        # Dashboard with advanced analytics and dynamic mini-graphs
        self.dashboard_label = ttk.Label(self.dashboard_tab, text="Dashboard", font=('Inter', 16, 'bold'))
        self.dashboard_label.pack(pady=10)

        # Frame pro výběr sloupců pro statistiky
        self.stats_options_frame = ttk.Frame(self.dashboard_tab)
        self.stats_options_frame.pack(fill=X, padx=10)

        self.stats_column_label = ttk.Label(self.stats_options_frame, text="Vyberte sloupec pro četnost hodnot:")
        self.stats_column_label.pack(side=LEFT, padx=5)

        self.stats_column_combobox = Combobox(self.stats_options_frame, state="readonly")
        self.stats_column_combobox.pack(side=LEFT, padx=5)

        self.stats_column_combobox.bind("<<ComboboxSelected>>", self.update_frequency_table)

        # Hlavní rámec pro dashboard
        self.dashboard_main_frame = ttk.Frame(self.dashboard_tab)
        self.dashboard_main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Horní rámec obsahující základní statistiky a korelační analýzu
        self.upper_frame = ttk.Frame(self.dashboard_main_frame)
        self.upper_frame.pack(fill=BOTH, expand=True)

        # Základní statistiky
        self.dashboard_stats_frame = LabelFrame(self.upper_frame, text="Základní statistiky")
        self.dashboard_stats_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        self.stats_tree = Treeview(self.dashboard_stats_frame, columns=("Metric", "Value"), show="headings")
        self.stats_tree.heading("Metric", text="Metrika")
        self.stats_tree.heading("Value", text="Hodnota")
        self.stats_tree.pack(fill=BOTH, expand=True)

        # Korelační analýza
        self.dashboard_correlation_frame = LabelFrame(self.upper_frame, text="Korelační analýza")
        self.dashboard_correlation_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        self.correlation_tree = Treeview(self.dashboard_correlation_frame, columns=("Variable Pair", "Correlation Coefficient"), show="headings")
        self.correlation_tree.heading("Variable Pair", text="Páry proměnných")
        self.correlation_tree.heading("Correlation Coefficient", text="Korelační koeficient")
        self.correlation_tree.pack(fill=BOTH, expand=True)

        # Dolní rámec obsahující četnost hodnot a dynamické grafy
        self.lower_frame = ttk.Frame(self.dashboard_main_frame)
        self.lower_frame.pack(fill=BOTH, expand=True)

        # Četnost hodnot
        self.frequency_frame = LabelFrame(self.lower_frame, text="Četnost hodnot")
        self.frequency_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        self.frequency_tree = Treeview(self.frequency_frame, columns=("Value", "Count"), show="headings")
        self.frequency_tree.heading("Value", text="Hodnota")
        self.frequency_tree.heading("Count", text="Počet výskytů")
        self.frequency_tree.pack(fill=BOTH, expand=True)

        # Dynamické grafy
        self.mini_graph_frame = LabelFrame(self.lower_frame, text="Grafy")
        self.mini_graph_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

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
                columns = list(self.data.columns)
                for column_combobox, _ in self.filters:
                    column_combobox['values'] = columns
                for column_combobox, _ in self.plot_filters:
                    column_combobox['values'] = columns
                self.plot_x_combobox['values'] = columns
                self.plot_y_combobox['values'] = columns
                self.stats_column_combobox['values'] = columns  # Přidáno pro výběr sloupce pro četnosti
                self.update_dashboard()
                messagebox.showinfo("Úspěch", "Soubor byl úspěšně načten!")
            except Exception as e:
                messagebox.showerror("Chyba", f"Soubor se nepodařilo načíst: {e}")

    def update_value_options(self, value_combobox, column_combobox):
        selected_column = column_combobox.get()
        if self.data is not None and selected_column in self.data.columns:
            unique_values = self.data[selected_column].dropna().unique()
            value_combobox['values'] = list(map(str, unique_values))

    def apply_filters(self):
        if self.data is not None:
            self.filtered_data = self.data
            for column_combobox, value_combobox in self.filters:
                column = column_combobox.get()
                value = value_combobox.get()
                if column and value:
                    self.filtered_data = self.filtered_data[self.filtered_data[column].astype(str) == value]
            messagebox.showinfo("Filtr", "Filtry byly aplikovány")
            self.update_table()
            self.update_dashboard()
        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte soubor.")

    def update_table(self):
        for col in self.tree.get_children():
            self.tree.delete(col)
        self.tree["column"] = list(self.filtered_data.columns)
        self.tree["show"] = "headings"
        for col in self.filtered_data.columns:
            self.tree.heading(col, text=col)
        for _, row in self.filtered_data.iterrows():
            self.tree.insert("", "end", values=list(row))

    def plot_data(self):
        if self.data is not None:
            # Použijeme filtrovaná data, pokud jsou dostupná, jinak originální data
            data_to_plot = self.filtered_data if self.filtered_data is not None else self.data

            # Aplikujeme filtry pro graf
            for column_combobox, value_combobox in self.plot_filters:
                column = column_combobox.get()
                value = value_combobox.get()
                if column and value:
                    data_to_plot = data_to_plot[data_to_plot[column].astype(str) == value]

            if not data_to_plot.empty:
                x_column = self.plot_x_combobox.get()
                y_column = self.plot_y_combobox.get()
                plot_type = self.plot_type_combobox.get()
                if x_column and y_column:
                    try:
                        plt.figure(figsize=(10, 6))
                        if plot_type == 'Scatter':
                            plt.scatter(data_to_plot[x_column], data_to_plot[y_column])
                        elif plot_type == 'Line':
                            plt.plot(data_to_plot[x_column], data_to_plot[y_column])
                        elif plot_type == 'Bar':
                            data_to_plot.groupby(x_column)[y_column].sum().plot(kind='bar')
                        elif plot_type == 'Histogram':
                            data_to_plot[y_column].plot(kind='hist')
                        else:
                            plt.scatter(data_to_plot[x_column], data_to_plot[y_column])

                        plt.title(f'{plot_type} graf {y_column} vs {x_column}')
                        plt.xlabel(x_column)
                        plt.ylabel(y_column)
                        plt.tight_layout()
                        plt.show()
                    except Exception as e:
                        messagebox.showerror("Chyba", f"Nepodařilo se vygenerovat graf: {e}")
                else:
                    messagebox.showwarning("Upozornění", "Vyberte prosím sloupce pro osu X a Y.")
            else:
                messagebox.showwarning("Upozornění", "Po aplikaci filtrů nejsou k dispozici žádná data.")
        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte data.")

    def update_dashboard(self):
        if self.data is not None:
            # Vymažeme předchozí obsah
            self.stats_tree.delete(*self.stats_tree.get_children())
            self.correlation_tree.delete(*self.correlation_tree.get_children())
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

            data_for_analysis = self.filtered_data if self.filtered_data is not None else self.data

            numeric_columns = data_for_analysis.select_dtypes(include="number").columns

            # Basic statistics for each numeric column
            for col in numeric_columns:
                mean = data_for_analysis[col].mean()
                median = data_for_analysis[col].median()
                std_dev = data_for_analysis[col].std()
                min_val = data_for_analysis[col].min()
                max_val = data_for_analysis[col].max()

                self.stats_tree.insert('', 'end', values=(f"{col} - Průměr", round(mean, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Medián", round(median, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Směrodatná odchylka", round(std_dev, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Minimum", round(min_val, 2)))
                self.stats_tree.insert('', 'end', values=(f"{col} - Maximum", round(max_val, 2)))

            # Correlation analysis between numeric columns
            for i in range(len(numeric_columns)):
                for j in range(i+1, len(numeric_columns)):
                    col1 = numeric_columns[i]
                    col2 = numeric_columns[j]
                    try:
                        corr_coef, _ = pearsonr(data_for_analysis[col1].dropna(), data_for_analysis[col2].dropna())
                        self.correlation_tree.insert('', 'end', values=(f"{col1} & {col2}", round(corr_coef, 2)))
                    except:
                        continue

            # Aktualizace četnostní tabulky
            selected_column = self.stats_column_combobox.get()
            if selected_column:
                self.update_frequency_table()

            # Přidání dynamických grafů
            self.add_dynamic_graphs(data_for_analysis)
        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte data.")

    def update_frequency_table(self, event=None):
        selected_column = self.stats_column_combobox.get()
        if selected_column and selected_column in self.data.columns:
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

            data_for_analysis = self.filtered_data if self.filtered_data is not None else self.data
            freq_series = data_for_analysis[selected_column].value_counts()

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
            canvas.get_tk_widget().pack(fill=BOTH, expand=True)
        else:
            self.frequency_tree.delete(*self.frequency_tree.get_children())
            for widget in self.mini_graph_frame.winfo_children():
                widget.destroy()

    def add_dynamic_graphs(self, data_for_analysis):
        # Vymažeme předchozí grafy
        for widget in self.mini_graph_frame.winfo_children():
            widget.destroy()

        # Přidáme graf četností, pokud je vybrán sloupec
        selected_column = self.stats_column_combobox.get()
        if selected_column and selected_column in data_for_analysis.columns:
            freq_series = data_for_analysis[selected_column].value_counts()
            fig, ax = plt.subplots(figsize=(5, 3))
            freq_series.plot(kind='bar', ax=ax)
            ax.set_title(f"Četnost hodnot ve sloupci {selected_column}")
            ax.set_xlabel("Hodnota")
            ax.set_ylabel("Počet výskytů")
            plt.xticks(rotation=45)
            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=self.mini_graph_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        # Přidáme další dynamické grafy
        if 'Datum Von' in data_for_analysis.columns and 'Fehler' in data_for_analysis.columns:
            fig2, ax2 = plt.subplots(figsize=(5, 3))
            # Convert dates to datetime format
            data_for_analysis['Datum Von'] = pd.to_datetime(data_for_analysis['Datum Von'], errors='coerce')
            # Group by date and count occurrences of "Fehler" per date
            data_for_analysis.set_index('Datum Von').resample('D')['Fehler'].count().plot(ax=ax2)
            ax2.set_title("Frekvence chyb v čase")
            ax2.set_xlabel("Datum")
            ax2.set_ylabel("Počet chyb")
            plt.tight_layout()
            canvas2 = FigureCanvasTkAgg(fig2, master=self.mini_graph_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=BOTH, expand=True)

        if 'PPlatz' in data_for_analysis.columns:
            fig3, ax3 = plt.subplots(figsize=(5, 3))
            # Count occurrences of each unique value in "PPlatz"
            data_for_analysis['PPlatz'].value_counts().plot(kind='bar', ax=ax3)
            ax3.set_title("Četnost hodnot v PPlatz")
            ax3.set_xlabel("PPlatz")
            ax3.set_ylabel("Četnost")
            plt.tight_layout()
            canvas3 = FigureCanvasTkAgg(fig3, master=self.mini_graph_frame)
            canvas3.draw()
            canvas3.get_tk_widget().pack(fill=BOTH, expand=True)

if __name__ == "__main__":
    root = ttk.Window(themename="flatly")  # Nastavíme téma aplikace
    app = ExcelAnalyzerApp(root)
    root.mainloop()
