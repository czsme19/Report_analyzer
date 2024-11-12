# gui.py
# Modul s GUI komponentami
# GUI je momentálně 

# gui.py

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import BOTH, X, LEFT, RIGHT, TOP, W, E, N, S, filedialog, messagebox
from tkinter.ttk import Combobox, Treeview, LabelFrame
import datetime
from data_processing import DataProcessor
from plotting import Plotter
from dashboard import Dashboard
from PIL import Image, ImageTk  # Přidáno pro práci s obrázky


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Analyzer")
        self.data_processor = DataProcessor()
        self.plotter = Plotter()
        self.dashboard = Dashboard()
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

        # Přidání loga (pokud máte logo, odkomentujte následující řádky a nahraďte 'logo.png' cestou k vašemu logu)
        # self.logo_image = tk.PhotoImage(file="logo.png")
        # self.logo_label = ttk.Label(self.welcome_frame, image=self.logo_image)
        # self.logo_label.pack(pady=10)

        # Dynamický uvítací text na základě denní doby
        current_hour = datetime.datetime.now().hour
        if current_hour < 12:
            greeting = "Dobré ráno"
        elif current_hour < 18:
            greeting = "Dobrý den"
        else:
            greeting = "Dobrý večer"

        self.title_label = ttk.Label(self.welcome_frame, text=f"{greeting}, vítejte v aplikaci Advanced Excel Analyzer", font=("Helvetica", 24, "bold"))
        self.title_label.pack(pady=10)

        # Popis aplikace
        self.description_label = ttk.Label(self.welcome_frame, text="Analyzujte svá Excel data snadno a rychle. Nahrajte soubor a začněte.", font=("Helvetica", 14))
        self.description_label.pack(pady=5)

        # Přidání instrukcí
        self.instructions_label = ttk.Label(self.welcome_frame, text="Klikněte na oblast níže nebo přetáhněte soubor pro načtení Excel souboru.", font=("Helvetica", 12))
        self.instructions_label.pack(pady=5)

        # Přidání drag-and-drop oblasti s použitím tk.Label místo ttk.Label
        self.drop_area = tk.Label(
            self.welcome_frame,
            text="Přetáhněte soubor sem nebo klikněte pro výběr",
            font=("Helvetica", 12),
            relief="ridge",
            borderwidth=2,
            width=50,
            height=10,
            anchor="center"
        )
        self.drop_area.pack(pady=20)
        self.drop_area.bind("<Button-1>", lambda e: self.load_file())
        # Implementace drag-and-drop může vyžadovat další knihovny nebo nastavení

        # Tip dne pro uživatele
        self.tip_label = ttk.Label(self.welcome_frame, text="Tip dne: Používejte filtry pro rychlé nalezení klíčových informací ve vašich datech.", font=("Helvetica", 12, "italic"), foreground="gray")
        self.tip_label.pack(pady=10)

        # Přidání ikony nápovědy
        self.help_button = ttk.Button(self.welcome_frame, text="?", width=3, command=self.show_help)
        self.help_button.place(x=10, y=10)

    def show_help(self):
        messagebox.showinfo("Nápověda", "Pro načtení souboru klikněte na oblast pro načtení nebo přetáhněte soubor do okna.")

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            # Zobrazení progress baru
            progress = ttk.Progressbar(self.welcome_frame, orient=HORIZONTAL, length=200, mode='indeterminate')
            progress.pack(pady=10)
            progress.start()
            self.root.update_idletasks()
            # Načtení souboru
            self.data_processor.load_file_direct(file_path)
            progress.stop()
            progress.destroy()
            self.data = self.data_processor.data
            if self.data is not None:
                columns = self.data_processor.get_columns()
                for column_combobox, _ in self.filters:
                    column_combobox['values'] = columns
                for column_combobox, _ in self.plot_filters:
                    column_combobox['values'] = columns
                self.plot_x_combobox['values'] = columns
                self.plot_y_combobox['values'] = columns
                self.stats_column_combobox['values'] = columns
                self.update_dashboard()

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
            columns = self.data.columns.tolist()
            column_combobox['values'] = columns

    def remove_filter(self):
        if self.filters:
            filter_widgets = self.filters.pop()
            filter_widgets[0].master.destroy()  # Zničíme celý řádek s filtry

    def apply_filters(self):
        filters = []
        for column_combobox, value_combobox in self.filters:
            column = column_combobox.get()
            value = value_combobox.get()
            filters.append((column, value))
        self.data_processor.apply_filters(filters)
        self.filtered_data = self.data_processor.filtered_data
        self.update_table()
        self.update_dashboard()

    def update_value_options(self, value_combobox, column_combobox):
        selected_column = column_combobox.get()
        if self.data is not None and selected_column in self.data.columns:
            unique_values = self.data[selected_column].dropna().unique()
            value_combobox['values'] = list(map(str, unique_values))

    def update_table(self):
        for col in self.tree.get_children():
            self.tree.delete(col)
        if self.filtered_data is not None:
            self.tree["column"] = list(self.filtered_data.columns)
            self.tree["show"] = "headings"
            for col in self.filtered_data.columns:
                self.tree.heading(col, text=col)
            for _, row in self.filtered_data.iterrows():
                self.tree.insert("", "end", values=list(row))

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

    def plot_data(self):
        if self.data is not None:
            data_to_plot = self.filtered_data if self.filtered_data is not None else self.data

            # Aplikujeme filtry pro graf
            plot_filters = []
            for column_combobox, value_combobox in self.plot_filters:
                column = column_combobox.get()
                value = value_combobox.get()
                if column and value:
                    plot_filters.append((column, value))

            x_column = self.plot_x_combobox.get()
            y_column = self.plot_y_combobox.get()
            plot_type = self.plot_type_combobox.get()

            self.plotter.plot_data(data_to_plot, x_column, y_column, plot_type, plot_filters)
        else:
            messagebox.showwarning("Upozornění", "Nejprve prosím načtěte data.")

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

        # Nastavení dashboardu
        self.dashboard.setup_dashboard(self.stats_tree, self.correlation_tree, self.frequency_tree, self.mini_graph_frame)

    def load_file(self):
        self.data_processor.load_file()
        self.data = self.data_processor.data
        if self.data is not None:
            columns = self.data_processor.get_columns()
            for column_combobox, _ in self.filters:
                column_combobox['values'] = columns
            for column_combobox, _ in self.plot_filters:
                column_combobox['values'] = columns
            self.plot_x_combobox['values'] = columns
            self.plot_y_combobox['values'] = columns
            self.stats_column_combobox['values'] = columns
            self.update_dashboard()

    def update_frequency_table(self, event=None):
        selected_column = self.stats_column_combobox.get()
        if selected_column:
            data_for_analysis = self.filtered_data if self.filtered_data is not None else self.data
            self.dashboard.update_frequency_table(selected_column, data_for_analysis)

    def update_dashboard(self):
        data_for_analysis = self.filtered_data if self.filtered_data is not None else self.data
        selected_column = self.stats_column_combobox.get()
        self.dashboard.update_dashboard(data_for_analysis, selected_column)
