import time
import pandas as pd  # Usa pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
 
# ==================================================================
# CONFIGURAÇÕES (JÁ PREENCHIDAS POR VOCÊ)
# ==================================================================
URL_LOGIN = "https://www.valeshop.com.br/portal/valeshop/!AP_LOGIN?p_tipo=1"
 
LOCATORS = {
    'cliente_codigo': (By.NAME, 'p_id_cliente'),
    'contrato_informe': (By.NAME, 'p_nr_contrato'),
    'motorista_nome': (By.NAME, 'p_nm_motorista'),
    'motorista_matricula': (By.NAME, 'p_nr_matricula'),
    'veiculo_placa': (By.NAME, 'p_nr_placa_veiculo'),
    'veiculo_marca': (By.NAME, 'p_ds_marca_veiculo'),
    'veiculo_modelo': (By.NAME, 'p_nm_modelo_veiculo'),
}
# ==================================================================
 
 
def extrair_dados_planilha(caminho_arquivo, logger):
    """
    Lê os dados de células específicas da planilha (modelo PMDF) com PANDAS.
    """
    logger(f"Lendo planilha com pandas: {caminho_arquivo}")
    try:
        # Carrega a planilha, sem cabeçalho
        with open(caminho_arquivo, 'rb') as f:
            df = pd.read_excel(f, header=None, engine='openpyxl')        
           
       
        dados = {}
 
        # Mapeamento com base no layout (índice 0 do pandas):
        # D2 -> iloc[linha=1, coluna=3] (Linha 2, Coluna D)
        dados['nome'] = df.iloc[3, 2]
       
        # E4 -> iloc[linha=3, coluna=4] (Linha 4, Coluna E)
        matricula_bruta = df.iloc[4, 2]
       
        # D7 -> iloc[linha=6, coluna=3] (Linha 7, Coluna D)
        dados['placa'] = df.iloc[7, 2]
 
        # D6 -> iloc[linha=5, coluna=3] (Linha 6, Coluna D)
        veiculo_completo = str(df.iloc[6, 2])
 
        # Lógica de divisão da Marca/Modelo
        if veiculo_completo and ' ' in veiculo_completo and veiculo_completo.lower() != 'nan':
            partes = veiculo_completo.split(' ', 1)
            dados['marca'] = partes[0]
            dados['modelo'] = partes[1]
        else:
            # Se não tiver espaço, ou for "nan", usa o valor inteiro como marca
            dados['marca'] = veiculo_completo if veiculo_completo.lower() != 'nan' else ""
            dados['modelo'] = ""
           
        # Tratamento da Matrícula (para remover .0 se for lida como float)
        if pd.isna(matricula_bruta):
             dados['matricula'] = None # Trata se a célula estiver vazia
        else:
            # Converte para int (para remover .0) e depois para string
            dados['matricula'] = str(int(float(matricula_bruta)))
   
        # Validação simples
        if not all([dados['nome'], dados['matricula'], dados['placa']]): # Marca não é mais obrigatória
            logger(f"AVISO: Dados faltando na planilha {caminho_arquivo}.")
            logger(f"Verifique as células: D2 (Nome), E4 (Matrícula), D7 (Placa).")
            return None
           
        logger("Dados extraídos com sucesso (usando pandas).")
        logger(f"Nome: {dados['nome']}, Matrícula: {dados['matricula']}, Placa: {dados['placa']}")
        return dados
 
    except Exception as e:
        logger(f"ERRO ao ler a planilha com pandas '{caminho_arquivo}': {e}")
        logger("Verifique se o arquivo é uma planilha válida e se as células D2, E4, D6, D7 existem.")
        return None
 
def preencher_formulario_web(dados_veiculo, logger):
    """
    Inicia o Selenium, pausa para login manual e preenche o formulário.
    """
    logger("Iniciando automação com Selenium...")
   
    driver = None
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(URL_LOGIN)
       
        logger("="*50)
        logger("--- AÇÃO MANUAL NECESSÁRIA ---")
        logger("O navegador foi aberto.")
        logger("1. Faça o login no sistema.")
        logger("2. Resolva o CAPTCHA (se houver).")
        logger("3. Navegue até a tela 'Cadastrar Vendas sem Cartão'.")
        logger("\nIMPORTANTE: Quando o formulário (imagem 1) estiver VISÍVEL,")
        logger("pressione a tecla 'Enter' no CONSOLE (terminal) onde")
        logger("você iniciou a aplicação.")
        logger("="*50)
       
        input("Pressione Enter no CONSOLE (terminal) para continuar...")
       
        logger("Continuando automação... Preenchendo formulário.")
       
        wait = WebDriverWait(driver, 30)
        wait.until(EC.visibility_of_element_located(LOCATORS['cliente_codigo']))
       
        logger("Preenchendo campos fixos...")
        driver.find_element(*LOCATORS['cliente_codigo']).send_keys("3359")
        driver.find_element(*LOCATORS['contrato_informe']).send_keys("00101033590125")
       
        logger("Preenchendo dados do motorista...")
        driver.find_element(*LOCATORS['motorista_nome']).send_keys(dados_veiculo['nome'])
        driver.find_element(*LOCATORS['motorista_matricula']).send_keys(dados_veiculo['matricula'])
       
        logger("Preenchendo dados do veículo...")
        driver.find_element(*LOCATORS['veiculo_placa']).send_keys(dados_veiculo['placa'])
        driver.find_element(*LOCATORS['veiculo_marca']).send_keys(dados_veiculo['marca'])
        driver.find_element(*LOCATORS['veiculo_modelo']).send_keys(dados_veiculo['modelo'])
       
        logger("Formulário preenchido com sucesso!")
        logger("A automação irá pausar por 15 segundos antes de fechar o navegador.")
        time.sleep(15)
       
    except Exception as e:
        logger(f"ERRO durante a automação Selenium: {e}")
        logger("Verifique se os 'LOCATORS' no arquivo 'backend/veiculo.py' estão corretos.")
       
    finally:
        if driver:
            driver.quit()
        logger("Navegador fechado. Automação concluída.")