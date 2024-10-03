import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

class ExcelImporter:
    def __init__(self, root):
        self.root = root
        self.root.title("MOTOGRAU")
        self.root.configure(bg="#f0f0f0")

        style = ttk.Style()
        style.configure("TLabel", background="#f0f0f0", font=("Helvetica", 12))
        style.configure("TButton", font=("Helvetica", 12), padding=5)
        style.configure("TEntry", font=("Helvetica", 12), padding=5)
        style.configure("TText", font=("Helvetica", 12), padding=5)

        self.label_motoboy = ttk.Label(root, text="Nome do Motoboy:")
        self.label_motoboy.pack(pady=10)

        self.entry_motoboy = ttk.Entry(root, width=30)
        self.entry_motoboy.pack(pady=10)

        self.import_button = ttk.Button(root, text="Importar Planilha Excel", command=self.import_excel, width=20)
        self.import_button.pack(pady=10)

        self.sum_button = ttk.Button(root, text="Somar Taxas da Planilha", command=self.somar_taxas, width=20)
        self.sum_button.pack(pady=10)

        self.text_area = tk.Text(root, wrap='word', width=50, height=10, bg="#ffffff", borderwidth=2, relief="groove")
        self.text_area.pack(pady=10)

        self.label_creator = ttk.Label(root, text="MotoGrau ®", font=("Helvetica", 14, "bold"))
        self.label_creator.pack(pady=5)

        self.label_author = ttk.Label(root, text="Criado por João Lucas Estrela", font=("Helvetica", 10))
        self.label_author.pack(pady=5)

        self.taxas = self.load_taxas()
        self.df_imported = pd.DataFrame(columns=['Bairro', 'Taxa'])

    def load_taxas(self):
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            taxas_path = os.path.join(base_path, 'taxas.csv')
            df = pd.read_csv(taxas_path)
            df['Taxa'] = pd.to_numeric(df['Taxa'], errors='coerce')
            df['Bairro'] = df['Bairro'].str.lower()
            return df
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar taxas: {e}")
            return pd.DataFrame(columns=['Bairro', 'Taxa'])

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                df_motoboys = pd.read_excel(file_path)

                if 'bairro' not in df_motoboys.columns and 'bairros' not in df_motoboys.columns:
                    messagebox.showerror("Erro", "A planilha deve conter a coluna 'bairro' ou 'bairros'.")
                    return

                bairro_coluna = 'bairro' if 'bairro' in df_motoboys.columns else 'bairros'
                df_motoboys['Taxa'] = None
                df_motoboys[bairro_coluna] = df_motoboys[bairro_coluna].str.lower()

                for index, row in df_motoboys.iterrows():
                    bairro = row[bairro_coluna]
                    taxa = self.taxas.loc[self.taxas['Bairro'] == bairro, 'Taxa']
                    if not taxa.empty:
                        df_motoboys.at[index, 'Taxa'] = taxa.values[0]

                self.df_imported = df_motoboys
                display_df = self.df_imported[['bairro', 'Taxa']]
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(tk.END, display_df.to_string(index=False))

            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao importar a planilha: {e}")

    def somar_taxas(self):
        try:
            total_imported = self.df_imported['Taxa'].sum()
            motoboy = self.entry_motoboy.get()
            entregas_count = self.df_imported['Taxa'].count()
            resumo_bairros = self.contar_bairros()
            self.show_results(motoboy, total_imported, entregas_count, resumo_bairros)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao somar as taxas: {e}")

    def contar_bairros(self):
        # Remove o nome 'bairro' e capitaliza os bairros
        counts = self.df_imported['bairro'].value_counts()
        return "\n".join(f"{bairro.capitalize()}: {count}" for bairro, count in counts.items())

    def show_results(self, motoboy, total, entregas, resumo_bairros):
        result_window = tk.Toplevel(self.root)
        result_window.title("Resultados")
        result_window.geometry("350x250")

        result_text = (
            f"Soma total para {motoboy}: R$ {total:.2f}\n"
            f"Quantidade de entregas: {entregas}\n\n"
            f"Resumo de Bairros:\n{resumo_bairros}"
        )
        text_area = tk.Text(result_window, wrap='word', height=10, width=40)
        text_area.insert(tk.END, result_text)
        text_area.config(state=tk.NORMAL)  # Permite selecionar o texto
        text_area.pack(pady=10)

        # Botão para fechar a janela
        close_button = ttk.Button(result_window, text="Fechar", command=result_window.destroy)
        close_button.pack(pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelImporter(root)
    root.mainloop()
