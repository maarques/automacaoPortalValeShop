import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from classes.filesFunctions import FilesFunctions # Você mencionou esta classe, então a mantive
import threading
import time

# Importa o seu backend
import backend.veiculo as backend 

# Precisamos das "Expected Conditions" (EC) para esperar o botão
from selenium.webdriver.support import expected_conditions as EC
# (ActionChains não é mais necessário aqui)


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

        # Assumindo que FilesFunctions lida com o self.lbl_entrada_status
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
        arquivo = filedialog.askopenfilename(title="Selecione o arquivo de cadastro", filetypes=tipos_arquivo)
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

    def _callback_pausa_login_gui(self):
        """
        Esta função é chamada pelo backend para pausar a automação
        e esperar o login manual.
        """
        self.log("="*50)
        self.log("--- AÇÃO MANUAL NECESSÁRIA ---")
        self.log("O navegador foi aberto.")
        self.log("1. Faça o login no sistema.")
        self.log("2. Resolva o CAPTCHA (se houver).")
        self.log("="*50)
        
        messagebox.showinfo(
            "Pausa para Login",
            "O navegador foi aberto.\n\n"
            "Faça o login e resolva o CAPTCHA (se houver).\n\n"
            "Clique em 'OK' NESTA JANELA para o robô continuar."
        )
        self.log("Usuário clicou em OK. Continuando automação...")

    def processar_em_thread(self):
        """
        Lógica de submissão revertida para o "Plano D" (submeter o form),
        que é a mais robusta para este caso.
        """
        driver = None 
        try:
            caminho_arquivo = self.functions.arquivos_selecionados[0]
            self.log(f"Processando arquivo: {caminho_arquivo}")
            
            dados_cabecalho, lista_transacoes = backend.extrair_dados_planilha(caminho_arquivo, self.log)
            
            if not dados_cabecalho or not lista_transacoes:
                self.log("ERRO: Falha ao extrair dados (cabeçalho ou transações). Verifique o log.")
                self.frame.after(0, lambda: messagebox.showerror("Erro de Leitura", "Não foi possível ler os dados da planilha. Verifique o log."))
                return
            
            self.log(f"Dados de cabeçalho extraídos. {len(lista_transacoes)} transações encontradas.")

            driver, wait = backend.iniciar_e_logar(self.log, self._callback_pausa_login_gui)
            if not driver:
                raise Exception("Falha ao iniciar o navegador.")
            
            if not backend.navegar_ate_formulario(driver, wait, self.log):
                raise Exception("Falha ao navegar até o formulário de inclusão.")

            # 4. Inicia o LOOP de registros
            for i, transacao in enumerate(lista_transacoes):
                self.log(f"--- Processando Registro {i+1} de {len(lista_transacoes)} ---")
                self.log(f"Dados: {transacao}")
                
                dados_completos = {**dados_cabecalho, **transacao}
                
                # 5. Preenche os campos (agora com a lógica do popup)
                if not backend.preencher_um_registro(driver, wait, dados_completos, self.log):
                    self.log(f"ERRO ao preencher o registro {i+1}. Abortando.")
                    break 
                
                # 6. Submete o formulário
                self.log("Campos preenchidos. Submetendo o formulário principal...")
                try:
                    # --- ### LÓGICA DO PLANO D (SUBMIT FORM) ### ---
                    
                    # 1. Encontra o formulário principal
                    form_principal = wait.until(EC.presence_of_element_located(
                        backend.LOCATORS['form_principal']
                    ))
                    
                    # 2. Submete o formulário
                    form_principal.submit()
                    self.log("Evento 'submit' enviado ao FORM.")
                    # --- Fim da Lógica ---
                    
                    self.log("Aguardando 5 segundos após o submit...")
                    time.sleep(5) 

                except Exception as e_confirm:
                    self.log(f"ERRO: Não foi possível submeter o formulário principal. {e_confirm}")
                    raise e_confirm


                # 7. Prepara para o próximo registro
                if i < len(lista_transacoes) - 1:
                    self.log("Formulário enviado. Preparando para o próximo registro...")
            
            self.log("--- TODOS OS REGISTROS FORAM PROCESSADOS ---")
            self.log("Aguardando 10 segundos antes de fechar.")
            time.sleep(10)

        except Exception as e:
            self.log(f"\nERRO INESPERADO no processamento: {e}")
            self.frame.after(0, lambda: messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}"))
        finally:
            if driver:
                driver.quit()
                self.log("Navegador fechado.")
            self.frame.after(0, lambda: self.btn_gerar.config(text="REGISTRAR VEÍCULO", state='normal'))
            self.log("Processo finalizado.")
