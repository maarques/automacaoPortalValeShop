import tkinter as tk
from tkinter import ttk, messagebox
from classes.filesFunctions import FilesFunctions
import threading
import pandas as pd
import backend.veiculo as backend 


class AppCadastroVeiculo:
    def __init__(self, parent_tab):
        self.frame = ttk.Frame(parent_tab, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivo de Entrada (Excel)", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_gerar = ttk.Button(frame_acao, text="REGISTRAR VEÍCULO", command=self.iniciar_processamento_veiculo)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, height=10, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.functions = FilesFunctions(self.lbl_entrada_status, None, self.log_text)

        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivo (.xlsx)", command=self.selecionar_arquivo_veiculo)
        btn_arquivos.pack(fill=tk.X, pady=5)
        
        self.lbl_entrada_status.pack(fill=tk.X, pady=5) 

    def log(self, mensagem: str):
        if self.log_text:
            try:
                self.frame.after(0, self._log_update, mensagem)
            except tk.TclError:
                print(mensagem)
        else:
            print(mensagem)
            
    def _log_update(self, mensagem: str):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{mensagem}\n")
        self.log_text.see(tk.END) 
        self.log_text.config(state='disabled')
            
    def selecionar_arquivo_veiculo(self):
        tipos_arquivo = [
            ("Planilha Excel", "*.xlsx"),
            ("Todos os arquivos", "*.*")
        ]
        arquivo = tk.filedialog.askopenfilename(title="Selecione o arquivo de cadastro", filetypes=tipos_arquivo)
        if arquivo:
            self.functions.arquivos_selecionados = [arquivo] 
            if self.functions.lbl_entrada_status:
                self.functions.lbl_entrada_status.config(text=f"Arquivo selecionado: {arquivo}")
            self.log(f"Entrada definida: {arquivo}")

    def iniciar_processamento_veiculo(self):
        self.functions.limpar_log()
        self.log("--- Iniciando Registro de Veículo ---")
        
        if not self.functions.arquivos_selecionados:
            self.log("ERRO: Nenhum arquivo de entrada selecionado.")
            messagebox.showerror("Erro", "Nenhum arquivo de entrada selecionado.")
            return
            
        self.btn_gerar.config(text="PROCESSANDO...", state='disabled')
        
        threading.Thread(
            target=self.processar_em_thread,
            daemon=True
        ).start()

    def _funcao_de_pausa_para_login(self):
        self.log("="*50)
        self.log("--- AÇÃO MANUAL NECESSÁRIA ---")
        self.log("O navegador foi aberto.")
        self.log("1. Faça o login no sistema.")
        self.log("2. Resolva o CAPTCHA (se houver).")
        self.log("3. Navegue até a tela 'Cadastrar Vendas sem Cartão'.")
        self.log("="*50)
        
        messagebox.showinfo(
            "Ação Manual Necessária",
            "O navegador foi aberto.\n\n"
            "1. Faça o login no sistema.\n"
            "2. Resolva o CAPTCHA (se houver).\n"
            "3. Navegue até a tela 'Cadastrar Vendas sem Cartão'.\n\n"
            "Clique em 'OK' NESTA JANELA apenas quando o formulário estiver visível."
        )

    def processar_em_thread(self):
        try:
            caminho_arquivo = self.functions.arquivos_selecionados[0]
            self.log(f"Processando arquivo: {caminho_arquivo}")
            
            dados_veiculo = backend.extrair_dados_planilha(caminho_arquivo, self.log)
            
            if dados_veiculo:
                backend.preencher_formulario_web(
                    dados_veiculo, 
                    self.log, 
                    self._funcao_de_pausa_para_login
                )
            else:
                self.log("ERRO: Falha ao extrair dados da planilha. Verifique o log acima.")
                self.frame.after(0, lambda: messagebox.showerror("Erro de Leitura", "Não foi possível ler os dados da planilha. Verifique o log."))

        except Exception as e:
            self.log(f"\nERRO INESPERADO no processamento: {e}")
            self.frame.after(0, lambda: messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}"))
        finally:
            self.frame.after(0, lambda: self.btn_gerar.config(text="REGISTRAR VEÍCULO", state='normal'))


