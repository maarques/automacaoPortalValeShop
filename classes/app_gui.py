import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import automation.controller as controller

class AppGui:
    """
    Esta classe constrói a Interface Gráfica (GUI)
    e delega as ações para o 'automation.controller'.
    """
    def __init__(self, parent_tab):
        self.frame = ttk.Frame(parent_tab, padding=10)
        self.frame.pack(fill=tk.BOTH, expand=True)

        main_frame = ttk.Frame(self.frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Seção de Entrada ---
        frame_entrada = ttk.LabelFrame(main_frame, text="1. Selecionar Arquivo de Entrada (Excel)", padding=15)
        frame_entrada.pack(fill=tk.X, padx=10, pady=10)

        self.lbl_entrada_status = ttk.Label(frame_entrada, text="Nenhuma entrada selecionada.", wraplength=650)
        
        btn_arquivos = ttk.Button(frame_entrada, text="Selecionar Arquivo (.xlsx)", command=self.selecionar_arquivo_veiculo)
        btn_arquivos.pack(fill=tk.X, pady=5)
        self.lbl_entrada_status.pack(fill=tk.X, pady=5) 
        
        # --- Seção de Ação ---
        frame_acao = ttk.Frame(main_frame)
        frame_acao.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_gerar = ttk.Button(frame_acao, text="REGISTRAR VEÍCULO", command=self.iniciar_processamento_veiculo)
        self.btn_gerar.pack(fill=tk.X, ipady=10)

        # --- Seção de Log ---
        frame_log = ttk.LabelFrame(main_frame, text="Mensagens do Processo", padding=10)
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(frame_log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(frame_log, height=10, wrap=tk.WORD, yscrollcommand=scrollbar.set, state='disabled', font=('Courier New', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        # Lógica do 'filesFunctions' mesclada aqui
        self.caminho_arquivo = None

    def log(self, mensagem: str):
        """ Envia uma mensagem para a caixa de log (thread-safe). """
        if self.log_text:
            try:
                self.frame.after(0, self._log_update, mensagem)
            except tk.TclError:
                print(mensagem) # Fallback se a janela fechar
        else:
            print(mensagem)
            
    def _log_update(self, mensagem: str):
        """ Atualiza a caixa de texto do log (chamado por self.log). """
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{mensagem}\n")
        self.log_text.see(tk.END) 
        self.log_text.config(state='disabled')
            
    def limpar_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def selecionar_arquivo_veiculo(self):
        """ Abre a caixa de diálogo para selecionar o arquivo Excel. """
        tipos_arquivo = [("Planilha Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        arquivo = filedialog.askopenfilename(title="Selecione o arquivo de cadastro", filetypes=tipos_arquivo)
        
        if arquivo:
            self.caminho_arquivo = arquivo 
            self.lbl_entrada_status.config(text=f"Arquivo selecionado: {arquivo}")
            self.log(f"Entrada definida: {arquivo}")

    def _callback_pausa_login_gui(self):
        """ Mostra o popup de pausa para o login manual. """
        self.log("="*50)
        self.log("--- AÇÃO MANUAL NECESSÁRIA ---")
        self.log("O navegador foi aberto. Faça o login e resolva o CAPTCHA.")
        self.log("="*50)
        
        messagebox.showinfo(
            "Pausa para Login",
            "O navegador foi aberto.\n\n"
            "Faça o login e resolva o CAPTCHA (se houver).\n\n"
            "Clique em 'OK' NESTA JANELA para o robô continuar."
        )
        self.log("Usuário clicou em OK. Continuando automação...")
        
    def _callback_finalizacao(self, sucesso=True, erro=None):
        """ Reativa o botão e mostra a mensagem final (sucesso ou erro). """
        self.btn_gerar.config(text="REGISTRAR VEÍCULO", state='normal')
        if erro:
            self.log(f"\nERRO INESPERADO no processamento: {erro}")
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {erro}")
        elif sucesso:
            self.log("Processo finalizado com sucesso.")
            messagebox.showinfo("Sucesso", "Automação concluída com sucesso!")
            
    def iniciar_processamento_veiculo(self):
        # Valida o arquivo e inicia a thread de automação.
        self.limpar_log()
        self.log("--- Iniciando Registro de Veículo ---")
        
        if not self.caminho_arquivo:
            self.log("ERRO: Nenhum arquivo de entrada selecionado.")
            messagebox.showerror("Erro", "Nenhum arquivo de entrada selecionado.")
            return
            
        self.btn_gerar.config(text="PROCESSANDO...", state='disabled')
        
        # Inicia a thread, passando os callbacks (log, pausa, fim) para o controller
        threading.Thread(
            target=controller.run_automation_flow,
            args=(
                self.caminho_arquivo,
                self.log,
                self._callback_pausa_login_gui,
                lambda s, e: self.frame.after(0, self._callback_finalizacao, s, e)
            ),
            daemon=True
        ).start()
