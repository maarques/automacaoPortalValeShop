import time
import os
from dotenv import load_dotenv
import automation.core_functions as core
from selenium.webdriver.support import expected_conditions as EC

def run_automation_flow(caminho_arquivo, logger, callback_pausa, callback_final):
    """
    Função principal que orquestra todo o processo de automação.
    Executada em uma thread separada.
    """
    driver = None
    try:
        # Carrega as variáveis de ambiente do .env
        load_dotenv()
        url_login = os.getenv("URL_LOGIN")
        if not url_login:
            raise Exception("URL_LOGIN não encontrada no arquivo .env")
        
        # Extrai dados da planilha
        dados_cabecalho, lista_transacoes = core.extrair_dados_planilha(caminho_arquivo, logger)
        
        if not dados_cabecalho or not lista_transacoes:
            raise Exception("Falha ao extrair dados (cabeçalho ou transações). Verifique o log.")
        
        logger(f"{len(lista_transacoes)} transações encontradas. Iniciando navegador...")

        # Inicia o navegador e pausa para login
        driver, wait = core.iniciar_e_logar(url_login, logger, callback_pausa)
        if not driver:
            raise Exception("Falha ao iniciar o navegador.")
        
        # Navega até o formulário (APENAS UMA VEZ)
        if not core.navegar_ate_formulario(driver, wait, logger):
            raise Exception("Falha ao navegar até o formulário de inclusão.")

        # Inicia o LOOP de registros
        for i, transacao in enumerate(lista_transacoes):
            logger(f"--- Processando Registro {i+1} de {len(lista_transacoes)} ---")
            
            dados_completos = {**dados_cabecalho, **transacao}
            
            #  Preenche os campos
            if not core.preencher_um_registro(driver, wait, dados_completos, logger):
                raise Exception(f"Falha no preenchimento do registro {i+1}")
            
            # Submete o formulário
            logger("Campos preenchidos. Enviando formulário...")
            try:
                form_principal = wait.until(EC.presence_of_element_located(
                    core.LOCATORS['form_principal']
                ))
                form_principal.submit()
                logger("Formulário enviado.")
                
                # Espera para a página recarregar e limpar os campos
                logger("Aguardando 5 segundos...")
                time.sleep(5) 

            except Exception as e_confirm:
                logger(f"ERRO: Não foi possível submeter o formulário principal. {e_confirm}")
                raise e_confirm

            # Prepara para o próximo registro
            if i < len(lista_transacoes) - 1:
                logger("Formulário enviado. Preparando para o próximo registro...")
        
        logger("--- TODOS OS REGISTROS FORAM PROCESSADOS ---")
        logger("Aguardando 10 segundos antes de fechar.")
        time.sleep(10)
        
        callback_final(sucesso=True, erro=None)

    except Exception as e:
        callback_final(sucesso=False, erro=e)
        
    finally:
        if driver:
            driver.quit()
            logger("Navegador fechado.")
        logger("Thread de automação finalizada.")
