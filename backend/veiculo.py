import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys # Importa a classe Keys
from selenium.webdriver.common.action_chains import ActionChains
 
URL_LOGIN = "https://www.valeshop.com.br/portal/valeshop/!AP_LOGIN?p_tipo=1"
 
# --- Alterado de volta para By.NAME ---
LOCATORS = {
    'cliente_codigo': (By.NAME, 'p_cd_cliente'),
    'contrato_informe': (By.NAME, 'p_nr_contrato'),
    'motorista_nome': (By.NAME, 'p_nm_motorista'),
    'motorista_matricula': (By.NAME, 'p_nr_matricula'),
    'veiculo_placa': (By.NAME, 'p_nr_placa_veiculo'),
    'veiculo_marca': (By.NAME, 'p_ds_marca_veiculo'),
    'veiculo_modelo': (By.NAME, 'p_nm_modelo_veiculo'),
}

MAP_CHAVES_BUSCA = {
    'nome': ['CONDUTOR'],
    'matricula': ['CPF'],
    'destino': ['DESTINO'],
    'veiculo': ['VEÍCULO'],
    'placa': ['PLACA']
}
 
def _extrair_valor_busca(texto_celula):
    """Verifica se uma célula contém uma das chaves que procuramos."""
    if not isinstance(texto_celula, str):
        return None
    
    texto_limpo = texto_celula.strip().upper().rstrip(':')
    
    for chave_padrao, variacoes in MAP_CHAVES_BUSCA.items():
        if texto_limpo in variacoes:
            return chave_padrao
            
    return None 
 
 
def extrair_dados_planilha(caminho_arquivo, logger):
    """
    Lê a planilha inteira procurando pelas chaves e capturando
    o valor na célula à direita.
    """
    logger(f"Lendo planilha com pandas (modo busca): {caminho_arquivo}")
    try:
        with open(caminho_arquivo, 'rb') as f:
            df = pd.read_excel(f, header=None, engine='openpyxl', dtype=str)
        
        dados = {}
        chaves_encontradas = set()
        chaves_necessarias = set(MAP_CHAVES_BUSCA.keys())

        for row in df.itertuples(index=False, name=None):
            for i, cell_value in enumerate(row):
                
                chave_encontrada = _extrair_valor_busca(cell_value)
                
                if chave_encontrada and chave_encontrada not in chaves_encontradas:
                    valor = None
                    
                    if (i + 1) < len(row):
                        valor = row[i+1]
                        
                    if valor and not pd.isna(valor) and str(valor).strip():
                        logger(f"Encontrado '{chave_encontrada}': {valor}")
                        dados[chave_encontrada] = str(valor).strip()
                        chaves_encontradas.add(chave_encontrada)
                    else:
                        logger(f"AVISO: Chave '{chave_encontrada}' encontrada, mas valor à direita está vazio.")
                
                if chaves_encontradas == chaves_necessarias:
                    break
            if chaves_encontradas == chaves_necessarias:
                break
        
        veiculo_completo = dados.get('veiculo', '')
        if veiculo_completo and ' ' in veiculo_completo:
            partes = veiculo_completo.split(' ', 1)
            dados['marca'] = partes[0]
            dados['modelo'] = partes[1]
        else:
            dados['marca'] = veiculo_completo
            dados['modelo'] = "" 

        matricula_bruta = dados.get('matricula')
        if matricula_bruta:
            try:
                dados['matricula'] = str(int(float(matricula_bruta)))
            except ValueError:
                matricula_limpa = str(matricula_bruta).strip()
                matricula_limpa = ''.join(filter(str.isdigit, matricula_limpa))
                dados['matricula'] = matricula_limpa
        
        if not all(k in dados for k in ['nome', 'matricula', 'placa']):
            logger(f"AVISO: Dados essenciais faltando na planilha {caminho_arquivo}.")
            logger(f"Verifique se as chaves 'CONDUTOR', 'CPF' e 'PLACA' (e seus valores) existem.")
            return None
            
        logger("Dados extraídos com sucesso (modo busca).")
        logger(f"Nome: {dados.get('nome')}, Matrícula: {dados.get('matricula')}, Placa: {dados.get('placa')}, Veículo: {dados.get('veiculo')}, Destino: {dados.get('destino')}")
        return dados

    except Exception as e:
        logger(f"ERRO ao ler a planilha com pandas '{caminho_arquivo}': {e}")
        logger("Verifique se o arquivo é uma planilha válida.")
        return None
 
def preencher_formulario_web(dados_veiculo, logger, callback_pausa_login):
    """
    Inicia o Selenium, pausa para login, entra no frame
    e preenche o formulário.
    """
    logger("Iniciando automação com Selenium...")
    
    driver = None
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(URL_LOGIN)
        
        if callback_pausa_login:
            callback_pausa_login()
        else:
            logger("AVISO: Nenhum callback de pausa fornecido. Pausando no console.")
            input("Pressione Enter no CONSOLE (terminal) para continuar...")
        
        logger("Continuando automação... Preenchendo formulário.")
        
        wait = WebDriverWait(driver, 30)
        
        # 1. Mudar para o <frame> principal (Nível 1)
        try:
            logger("Tentando mudar para o frame 'content' (Nível 1)...")
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "content")))
            logger("Mudança para frame 'content' (Nível 1) bem-sucedida.")
            
        except Exception as e:
            logger(f"ERRO CRÍTICO: Não foi possível mudar para o frame 'content' (Nível 1). Erro: {e}")
            raise e
        
        try:
            logger("Localizando e clicando no botão 'CONTROLLER'...")
            # Usa o localizador mais simples e correto: O texto do link
            controller_link = wait.until(EC.element_to_be_clickable((
                By.LINK_TEXT, "CONTROLLER"
            )))
            controller_link.click()
            logger("Clique em 'CONTROLLER' efetuado.")
        except Exception as e:
            logger(f"ERRO: Não foi possível clicar em 'CONTROLLER'. {e}")
            logger("Verifique se o texto 'CONTROLLER' está exatamente assim (maiúsculas).")
            raise

        # 2. Mudar para o <frame> ANINHADO (Nível 2)
        try:
            logger("Tentando mudar para o frame 'content' (Nível 2, aninhado)...")
            # O driver agora procura *dentro* do Nível 1
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "content")))
            logger("Mudança para frame 'content' (Nível 2) bem-sucedida.")
            
        except Exception as e:
            logger(f"ERRO CRÍTICO: Não foi possível mudar para o frame 'content' aninhado (Nível 2). Erro: {e}")
            raise e

        try:
            logger("Navegando no menu (Veículo > Compras sem Cartão > Inclusão)...")

            # 1. Espera o menu "Veículo" aparecer
            menu_veiculo = wait.until(EC.visibility_of_element_located((
                By.LINK_TEXT, "Veículo"
            )))

            actions = ActionChains(driver)
            actions.move_to_element(menu_veiculo).perform()   # Paira o mouse sobre "Veículo"

            # 2. Localiza os submenus
            menu_compras = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Compras sem cartão")))
            menu_compras.click()  # Paira o mouse sobre "Compras sem Cartão"
            
            menu_inclusao = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Inclusão")))
            menu_inclusao.click()                      # Executa a ação
            
            logger("Navegação do menu concluída.")
        except Exception as e:
            logger(f"ERRO: Falha ao navegar no menu 'Veículo'. {e}")
            raise

        # 3. Preencher os campos (agora estamos no Nível 2)
        try:
            logger("Esperando pelo primeiro campo (Cliente) com NAME 'p_id_cliente'...")
            wait.until(EC.presence_of_element_located(LOCATORS['cliente_codigo']))
            logger("Primeiro campo encontrado.")
            
            # --- LÓGICA DE LIMPAR E PREENCHER (CTRL+A) ---
            campo_cliente = driver.find_element(*LOCATORS['cliente_codigo'])
            campo_cliente.send_keys(Keys.CONTROL, "a") # Seleciona tudo
            campo_cliente.send_keys("3359") # Sobrescreve
            logger("Preenchimento do primeiro campo bem-sucedido.")
            
        except NoSuchElementException as e:
            logger(f"ERRO: Falha ao encontrar o primeiro campo ({LOCATORS['cliente_codigo']}) mesmo após o switch duplo.")
            raise e # Relança o erro
        
        logger("Preenchendo campos fixos...")
        campo_contrato = driver.find_element(*LOCATORS['contrato_informe'])
        campo_contrato.send_keys(Keys.CONTROL, "a")
        campo_contrato.send_keys("00101033590125")
        
        logger("Preenchendo dados do motorista...")
        campo_nome = driver.find_element(*LOCATORS['motorista_nome'])
        campo_nome.send_keys(Keys.CONTROL, "a")
        campo_nome.send_keys(dados_veiculo['nome'])
        
        campo_matricula = driver.find_element(*LOCATORS['motorista_matricula'])
        campo_matricula.send_keys(Keys.CONTROL, "a")
        campo_matricula.send_keys(dados_veiculo['matricula'])
               
        logger("Preenchendo dados do veículo...")
        campo_placa = driver.find_element(*LOCATORS['veiculo_placa'])
        campo_placa.send_keys(Keys.CONTROL, "a")
        campo_placa.send_keys(dados_veiculo['placa'])
        
        campo_marca = driver.find_element(*LOCATORS['veiculo_marca'])
        campo_marca.send_keys(Keys.CONTROL, "a")
        campo_marca.send_keys(dados_veiculo['marca'])
        
        campo_modelo = driver.find_element(*LOCATORS['veiculo_modelo'])
        campo_modelo.send_keys(Keys.CONTROL, "a")
        campo_modelo.send_keys(dados_veiculo['modelo'])
        
        logger("Formulário preenchendido com sucesso!")
        logger("A automação irá pausar por 15 segundos antes de fechar o navegador.")
        time.sleep(15)
        
    except Exception as e:
        logger(f"ERRO durante a automação Selenium: {e}")
        logger("Verifique se os 'LOCATORS' (By.NAME) no arquivo 'backend/veiculo.py' estão corretos.")
        
    finally:
        if driver:
          logger("Navegador fechado. Automação concluída.")

