import tkinter as tk
from tkinter import ttk
# A linha de 'etiquetas' foi removida
from classes.veiculos import AppCadastroVeiculo

if __name__ == "__main__":
    root = tk.Tk()
    # Título focado na automação
    root.title("Automação de Registro de Veículos - Valeshop")
    root.geometry("800x700") 

    # Estilos globais (mantidos)
    style = ttk.Style()
    style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=10)
    style.configure('TLabel', font=('Helvetica', 10), padding=5)
    style.configure('TFrame', padding=10)
    style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))

    # --- Lógica de Abas (Notebook) Removida ---

    # Cria um frame principal que preenche toda a janela
    main_frame = ttk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Inicia o aplicativo de veículos diretamente dentro do frame principal
    # (A classe AppCadastroVeiculo espera um 'parent_tab', 
    # e o main_frame agora serve como esse 'pai')
    app_veiculos = AppCadastroVeiculo(main_frame)

    root.mainloop()
