import tkinter as tk
from tkinter import ttk
from classes.app_gui import AppGui

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Automação de Registro de Veículos - Valeshop")
    root.geometry("800x700") 

    # Estilos globais
    style = ttk.Style()
    style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=10)
    style.configure('TLabel', font=('Helvetica', 10), padding=5)
    style.configure('TFrame', padding=10)
    style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))

    main_frame = ttk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    app_veiculos = AppGui(main_frame)

    root.mainloop()