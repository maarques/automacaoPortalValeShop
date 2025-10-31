import time
import pandas as pd  # Usa pandas
import re  # Importa expressões regulares
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
 
# ==================================================================
# CONFIGURAÇÕES
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
 
# --- MAPEAMENTO DE CHAVES PARA BUSCA ---
# Define as chaves que buscamos na planilha (lado esquerdo)
# e o nome do campo que elas representarão (lado direito)
MAP_CHAVES_BUSCA = {
    'nome': ['CONDUTOR'],
    'matricula': ['CPF'],
    'destino': ['DESTINO'],
    'veiculo': ['VEÍCULO'], # Capturará a string completa (ex: "MiTSUBISHI 4454")
    'placa': ['PLACA']
}
# ==================================================================
 
def _extrair_valor_busca(texto_celula):
    """
    Função auxiliar para verificar se um texto de célula corresponde a uma chave.
    """
    if not isinstance(texto_celula, str):
        return None
    
    # Limpa o texto: "  CONDUTOR:  " -> "CONDUTOR"
    texto_limpo = texto_celula.strip().rstrip(':').strip()
    
    for chave_padrao, variacoes in MAP_CHAVES_BUSCA.items():
        for variacao in variacoes:
            # Compara ignorando maiúsculas/minúsculas
            if texto_limpo.lower() == variacao.lower():
                return chave_padrao # Retorna a chave (ex: 'nome')
    return None # Não é uma chave que procuramos
 
 
def extrair_dados_planilha(caminho_arquivo, logger):
    """
    Lê os dados da planilha iterando pelas células em busca de chaves.
    """
    logger(f"Lendo planilha com pandas: {caminho_arquivo}")
    try:
        # Carrega a planilha, sem cabeçalho
        with open(caminho_arquivo, 'rb') as f:
            df = pd.read_excel(f, header=None, engine='openpyxl')        
           
        dados = {}
        chaves_encontradas = set()
        chaves_necessarias = set(MAP_CHAVES_BUSCA.keys())
        
        # Itera por todas as linhas
        for row_idx, row in df.iterrows():
            # Itera por todas as células da linha
            for col_idx, celula in enumerate(row):
                
                chave_encontrada = _extrair_valor_busca(celula)
                
                # Se a célula é uma chave (ex: "CONDUTOR")
                if chave_encontrada:
                    # Pega o valor na célula à direita (col_idx + 1)
                    if (col_idx + 1) < len(row):
                        valor = row.iloc[col_idx + 1] # Pega o valor
                        
                        if not pd.isna(valor):
                            dados[chave_encontrada] = str(valor).strip()
                            chaves_encontradas.add(chave_encontrada)
            
            # Se já encontramos todas as chaves, paramos de ler a planilha
            if chaves_encontradas == chaves_necessarias:
                logger("Todas as chaves necessárias foram encontradas.")
                break 
 
        # --- PÓS-PROCESSAMENTO ---
        
        # Verifica se todas as chaves foram realmente encontradas
        if chaves_encontradas != chaves_necessarias:
            chaves_faltantes = chaves_necessarias - chaves_encontradas
            logger(f"ERRO: Não foi possível encontrar todas as chaves. Faltando: {chaves_faltantes}")
            logger(f"Verifique se a planilha contém os rótulos: {list(MAP_CHAVES_BUSCA.values())}")
            return None

        # Tratamento da Matrícula (CPF)
        if 'matricula' in dados:
            matricula_bruta = dados['matricula']
            try:
                # Tenta formato numérico
                dados['matricula'] = str(int(float(matricula_bruta)))
            except ValueError:
                # Se falhar, assume formato texto (ex: 810.881.321-20)
                matricula_limpa = str(matricula_bruta).strip()
                matricula_limpa = ''.join(filter(str.isdigit, matricula_limpa))
                dados['matricula'] = matricula_limpa
        
        # Lógica de divisão da Marca/Modelo (baseado na chave 'veiculo')
        veiculo_completo = dados.get('veiculo', '')
        if veiculo_completo and ' ' in veiculo_completo:
            partes = veiculo_completo.split(' ', 1)
            dados['marca'] = partes[0]
            dados['modelo'] = partes[1]
        else:
            dados['marca'] = veiculo_completo
            dados['modelo'] = ""
           
        # Validação final (campos obrigatórios para o formulário)
        if not all([dados.get('nome'), dados.get('matricula'), dados.get('placa')]):
            logger(f"AVISO: Dados essenciais (Nome, Matrícula ou Placa) estão faltando após a extração.")
            return None
           
        logger("Dados extraídos com sucesso (usando busca).")
        logger(f"Nome: {dados.get('nome')}, Matrícula: {dados.get('matricula')}, Placa: {dados.get('placa')}, Destino: {dados.get('destino')}, Veículo: {dados.get('marca')} {dados.get('modelo')}")
        return dados
 
    except Exception as e:
        logger(f"ERRO ao ler a planilha com pandas '{caminho_arquivo}': {e}")
        logger("Verifique se o arquivo é uma planilha válida.")
        return None
 
def preencher_formulario_web(dados_veiculo, logger, callback_pausa_login):
    """
    Inicia o Selenium, pausa para login manual e preenche o formulário.
    Usa um 'callback' para pausar em vez de 'input()'.
    """
    logger("Iniciando automação com Selenium...")
   
    driver = None
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(URL_LOGIN)
       
        # Chama a função de pausa (que mostrará o messagebox na UI)
        if callback_pausa_login:
            callback_pausa_login()
        else:
            # Fallback caso nenhum callback seja passado
            logger("AVISO: Nenhum callback de pausa fornecido. Pausando no console.")
            input("Pressione Enter no CONSOLE (terminal) para continuar...")
       
        logger("Continuando automação... Preenchendo formulário.")
       
        wait = WebDriverWait(driver, 30)
        # Espera o primeiro campo do formulário ficar visível
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
       
        logger("Formulário preenchendido com sucesso!")
        logger("A automação irá pausar por 15 segundos antes de fechar o navegador.")
        time.sleep(15)
       
    except Exception as e:
        logger(f"ERRO durante a automação Selenium: {e}")
        logger("Verifique se os 'LOCATORS' no arquivo 'backend/veiculo.py' estão corretos.")
       
    finally:
        if driver:
            driver.quit()
        logger("Navegador fechado. Automação concluída.")

