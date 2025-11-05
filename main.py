import tkinter as tk
from tkinter import ttk
from classes.etiquetas import AppGeradorEtiquetas   
from classes.veiculos import AppCadastroVeiculo

class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador Cebraspe")
        self.root.geometry("800x700") 

        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=10)
        style.configure('TLabel', font=('Helvetica', 10), padding=5)
        style.configure('TFrame', padding=10)
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))
        style.configure('TNotebook.Tab', font=('Helvetica', 10, 'bold'), padding=[10, 5])

        notebook = ttk.Notebook(root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_etiquetas = ttk.Frame(notebook)
        notebook.add(self.tab_etiquetas, text="Gerar Etiquetas")
        self.app_etiquetas = AppGeradorEtiquetas(self.tab_etiquetas) 

        self.tab_veiculos = ttk.Frame(notebook)
        notebook.add(self.tab_veiculos, text="Cadastrar Veículo")
        self.app_veiculos = AppCadastroVeiculo(self.tab_veiculos)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApplication(root)
    root.mainloop()
