import tkinter as tk
import backend.veiculo as backend
from tkinter import ttk
from classes.filesFunctions import FilesFunctions

class AppCadastroVeiculo:
    def __init__(self, parent_tab):
        self.frame = ttk.Frame(parent_tab, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivos de Entrada", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_gerar = ttk.Button(frame_acao, text="REGISTRAR VEÍCULO", command=self.registrar_veiculo_placeholder)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, height=10, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.functions = FilesFunctions(self.lbl_entrada_status, None, self.log_text)

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivos (.docx, .xlsx)", command=self.functions.selecionar_arquivos)
        btn_arquivos.pack(fill=tk.X, pady=5)

        # btn_pasta = ttk.Button(frame_entrada, text="Selecionar Pasta (contém .docx, .xlsx)", command=self.functions.selecionar_pasta)
        # btn_pasta.pack(fill=tk.X, pady=5)
        
        self.lbl_entrada_status.pack(fill=tk.X, pady=5) 


    def log(self, mensagem: str):
        if self.log_text:
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{mensagem}\n")
            self.log_text.see(tk.END) # Auto-scroll
            self.log_text.config(state='disabled')
        else:
            print(mensagem)
            
    def registrar_veiculo_placeholder(self):
        self.functions.limpar_log()
        self.log("--- Iniciando Registro de Veículo ---")
        
        if not self.functions.arquivos_selecionados:
            self.log("ERRO: Nenhum arquivo de entrada selecionado.")
            return
            
        self.log(f"Arquivos/Pastas selecionados: {self.functions.arquivos_selecionados}")
        self.log("TODO: Implementar lógica de registro de veículo.")
