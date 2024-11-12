# plotting.py
# Modul pro vytváření grafů

import matplotlib.pyplot as plt
from tkinter import messagebox

class Plotter:
    def plot_data(self, data, x_column, y_column, plot_type, plot_filters):
        if data is not None:
            data_to_plot = data

            # Aplikujeme filtry pro graf
            for column, value in plot_filters:
                if column and value:
                    data_to_plot = data_to_plot[data_to_plot[column].astype(str) == value]

            if not data_to_plot.empty:
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
