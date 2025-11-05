import time
import pandas as pd
import re
import math
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select

MAP_CHAVES_BUSCA_CABECALHO = {
    'nome': ['CONDUTOR'],
    'matricula': ['CPF'],
    'destino': ['DESTINO'],
    'veiculo': ['VEÍCULO'],
    'placa': ['PLACA'],
}

LOCATORS = {
    'cliente_codigo': (By.NAME, 'p_cd_cliente'),
    'contrato_informe': (By.NAME, 'p_nr_contrato'),
    'motorista_nome': (By.NAME, 'p_nm_motorista'),
    'motorista_matricula': (By.NAME, 'p_nr_matricula'),
    'hodometro_abastecimento': (By.NAME, 'p_nr_km_veiculo'), 
    'destino': (By.NAME, 'p_nm_razao_social_cre'),
    'data': (By.NAME, 'p_dt_especifica'),
    'select_produto': (By.NAME, 'p_id_produto_servico'),
    'litros': (By.NAME, 'p_qt_produto_utilizado'),
    'valor_total': (By.NAME, 'p_vl_produto'),
    'btn_incluir_novo': (By.XPATH, "//button[contains(text(), 'Incluir novo')]"),
    'hora': (By.NAME, 'p_hr_especifica'),
    'btn_confirmar': (By.CSS_SELECTOR, "button.gravar"),
    'form_principal': (By.XPATH, "//button[@class='gravar']/ancestor::form")
}

# --- Lógica do Popup da Placa ---

def _buscar_placa_popup(driver, placa, logger):
    wait = WebDriverWait(driver, 20)
 
    try:
        # Clica no link para abrir o modal
        link_buscar_locator = (By.CSS_SELECTOR, "a.jqModal.buscar[href*='!lv_placa']")
        link_buscar = wait.until(EC.element_to_be_clickable(link_buscar_locator))
        driver.execute_script("arguments[0].scrollIntoView(true);", link_buscar)
        link_buscar.click()
 
        # Espera o modal ficar visível
        modal_locator = (By.CSS_SELECTOR, "#lov.jqmWindow")
        wait.until(EC.visibility_of_element_located(modal_locator))
 
        # Espera pelo campo de input (ID duplicado)
        input_placa_locator = (By.ID, "p_nr_placa_veiculo")
        wait.until(
            lambda d: len(d.find_elements(*input_placa_locator)) > 1,
            "Modal abriu, mas o campo de input duplicado não foi encontrado."
        )
        
        # Pega o segundo campo [1], que é o do modal
        input_placa = driver.find_elements(*input_placa_locator)[1]
        
        input_placa.clear()
        input_placa.send_keys(placa)
 
        # Força a validação do campo (onBlur/onChange)
        input_placa.send_keys(Keys.TAB)
        time.sleep(0.5)

        # Clica no botão "Pesquisar" (dentro do modal)
        botao_pesquisar_locator = (By.XPATH, "//button[contains(text(), 'Pesquisar')]")
        botao_pesquisar = wait.until(EC.element_to_be_clickable(botao_pesquisar_locator))
        botao_pesquisar.click()
 
        # Espera a tabela de resultados
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.tabela")))
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "a.retorno")) > 0, 
                   "A busca no popup não retornou nenhum resultado.")
        time.sleep(0.5)  
 
        # Tenta clicar no resultado exato, se falhar, clica no primeiro
        try:
            link_resultado_locator = (By.XPATH, f"//a[@class='retorno' and starts-with(normalize-space(text()), '{placa.upper()}')]")
            link_resultado = wait.until(EC.element_to_be_clickable(link_resultado_locator))
        except:
            logger("AVISO: Placa exata não encontrada no popup. Clicando no primeiro resultado (fallback).")
            link_resultado_locator = (By.CSS_SELECTOR, "a.retorno")
            link_resultado = wait.until(EC.element_to_be_clickable(link_resultado_locator))
 
        link_resultado.click()
 
        # Espera o modal ficar invisível
        wait.until(EC.invisibility_of_element_located((By.ID, "lov")))
 
        # Verificação final
        campo_placa_principal = wait.until(EC.presence_of_element_located((By.NAME, "p_nr_placa_veiculo")))
        placa_final = campo_placa_principal.get_attribute("value")
        logger(f"Placa '{placa_final}' preenchida via popup.")
        
        return True

    except Exception as e:
        logger(f"ERRO CRÍTICO durante busca da placa no popup: {e}")
        raise e 
 
# --- Funções de Extração ---

def _extrair_valor_busca(texto_celula, mapa_busca):
    if not isinstance(texto_celula, str):
        return None
    texto_limpo = texto_celula.strip().upper().rstrip(':')
    for chave_padrao, variacoes in mapa_busca.items():
        if texto_limpo in variacoes:
            return chave_padrao
    return None

def _mapear_colunas_tabela(header_row, logger):
    mapa = {}
    for i, cell in enumerate(header_row):
        val = str(cell).strip().upper()
        if val == 'DATA':
            mapa['data'] = i
        elif val == 'HODÔMETRO ABASTECIMENTO':
            mapa['hodometro_abastecimento'] = i
        elif val == 'COMBUSTÍVEL':
            mapa['produto_nome'] = i
        elif 'VALOR' in val:
            mapa['valor_total'] = i
        elif val == 'LITROS':
            mapa['litros'] = i
    logger(f"Mapa de colunas da tabela criado: {mapa}")
    return mapa

def _extrair_transacao(row_data, col_map, logger):
    try:
        product_regex = re.compile(
            r"(.*?)\s*\(\s*([\d,.]+)\s*\)\s*\+\s*(Aditivo|Arla)\s*\(\s*([\d,.]+)\s*\)", 
            re.IGNORECASE
        )
        idx_data = col_map.get('data', 0)
        val_data = str(row_data[idx_data]).strip()
        idx_check_total_col = col_map.get('valor_total', 4)
        val_check_total_cell = str(row_data[idx_check_total_col - 1]).strip().upper()

        if pd.isna(val_data) or val_data == "" or "TOTAL" in val_check_total_cell:
            return None

        data_obj = pd.to_datetime(val_data)
        data_formatada = data_obj.strftime('%d/%m/%Y')
        
        produto_nome_bruto = str(row_data[col_map['produto_nome']])
        match = product_regex.search(produto_nome_bruto)
        
        aditivo_info = None
        produto_nome_final = produto_nome_bruto.strip()
        valor_total_final = str(row_data[col_map['valor_total']]).strip().replace(',', '.')
        
        if match:
            produto_nome_final = match.group(1).strip()
            valor_diesel = match.group(2).strip().replace(',', '.')
            nome_aditivo = match.group(3).strip()
            valor_aditivo_bruto = match.group(4).strip().replace(',', '.')
            valor_total_final = valor_diesel
            
            try:
                valor_aditivo_float = float(valor_aditivo_bruto)
                litros_base = math.ceil(valor_aditivo_float / 40)
                litros_aditivo_formatado = str(int(litros_base)) + "00" # Formato "100"
                
                aditivo_info = {
                    'valor': valor_aditivo_bruto,
                    'litros': litros_aditivo_formatado
                }
                logger(f"Produto dividido: {produto_nome_final} ({valor_diesel}) + {nome_aditivo} ({valor_aditivo_bruto})")
            except Exception as e_calc:
                logger(f"ERRO ao calcular litros do aditivo para valor '{valor_aditivo_bruto}': {e_calc}")
        
        idx_litros = col_map.get('litros', 5)
        litros_diesel = str(row_data[idx_litros]).replace('L', '').replace('"', '').strip().replace(',', '.')
        
        transacao = {
            'data': data_formatada,
            'hodometro_abastecimento': str(int(float(row_data[col_map['hodometro_abastecimento']]))),
            'produto_nome': produto_nome_final,
            'valor_total': valor_total_final,
            'litros': litros_diesel,
            'aditivo': aditivo_info
        }
        return transacao
    except Exception as e:
        logger(f"AVISO: Ignorando linha (provavelmente cabeçalho/total ou erro). Linha: {row_data}. Erro: {e}")
        return None

def extrair_dados_planilha(caminho_arquivo, logger):
    logger(f"Lendo planilha: {caminho_arquivo}...")
    try:
        with open(caminho_arquivo, 'rb') as f:
            df = pd.read_excel(f, header=None, engine='openpyxl', dtype=str)
        
        dados_cabecalho = {}
        lista_transacoes = []
        chaves_cabecalho_encontradas = set()
        chaves_cabecalho_necessarias = set(MAP_CHAVES_BUSCA_CABECALHO.keys())
        idx_tabela_header = -1
        col_map = {}

        for idx, row in enumerate(df.itertuples(index=False, name=None)):
            if idx_tabela_header == -1:
                is_row_tabela_header = False
                for i, cell_value in enumerate(row):
                    chave_encontrada = _extrair_valor_busca(cell_value, MAP_CHAVES_BUSCA_CABECALHO)
                    if chave_encontrada and chave_encontrada not in chaves_cabecalho_encontradas:
                        valor = row[i+1] if (i + 1) < len(row) else None
                        if valor and not pd.isna(valor) and str(valor).strip():
                            logger(f"Dado encontrado: '{chave_encontrada}' = {valor}")
                            dados_cabecalho[chave_encontrada] = str(valor).strip()
                            chaves_cabecalho_encontradas.add(chave_encontrada)
                    
                    if str(cell_value).strip().upper() == 'DATA':
                        idx_tabela_header = idx
                        col_map = _mapear_colunas_tabela(row, logger)
                        is_row_tabela_header = True
                        break
                if is_row_tabela_header:
                    continue
            else:
                transacao = _extrair_transacao(row, col_map, logger)
                if transacao:
                    lista_transacoes.append(transacao)
                elif col_map and 'TOTAL' in str(row[col_map.get('valor_total', 4) - 1]).strip().upper():
                    logger("Fim da extração de transações (linha 'TOTAL' encontrada).")
                    break
            
        veiculo_completo = dados_cabecalho.get('veiculo', '')
        if veiculo_completo and ' ' in veiculo_completo:
            partes = veiculo_completo.split(' ', 1)
            dados_cabecalho['marca'] = partes[0]
            dados_cabecalho['modelo'] = partes[1]
        else:
            dados_cabecalho['marca'] = veiculo_completo
            dados_cabecalho['modelo'] = "" 
        
        matricula_bruta = dados_cabecalho.get('matricula')
        if matricula_bruta:
            try:
                dados_cabecalho['matricula'] = str(int(float(matricula_bruta)))
            except ValueError:
                dados_cabecalho['matricula'] = ''.join(filter(str.isdigit, str(matricula_bruta)))
        
        if not all(k in dados_cabecalho for k in ['nome', 'matricula', 'placa']):
            logger("ERRO: Dados essenciais do cabeçalho (Nome, Matrícula, Placa) não encontrados.")
            return None, []
        if not lista_transacoes:
            logger("ERRO: Nenhum registro de transação encontrado na tabela.")
            return dados_cabecalho, []
            
        logger("Injetando horas incrementais nas transações...")
        hora_inicial = datetime.now().hour
        for i, transacao in enumerate(lista_transacoes):
            hora_calculada = (hora_inicial + i) % 24
            hora_formatada = f"{hora_calculada:02d}00" # Formato HH00
            transacao['hora_para_preencher'] = hora_formatada

        logger(f"Extração concluída. Transações: {len(lista_transacoes)}")
        return dados_cabecalho, lista_transacoes
        
    except Exception as e:
        logger(f"ERRO CRÍTICO ao ler a planilha '{caminho_arquivo}': {e}")
        return None, []

# --- Funções de Navegação ---

def iniciar_e_logar(url, logger, callback_pausa_login):
    logger("Iniciando automação com Selenium...")
    driver = None
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.get(url)
        
        callback_pausa_login() 
        
        logger("Continuando automação...")
        wait = WebDriverWait(driver, 30)
        return driver, wait
    except Exception as e:
        logger(f"ERRO ao iniciar o navegador ou logar: {e}")
        if driver:
            driver.quit()
        return None, None

def navegar_ate_formulario(driver, wait, logger):
    try:
        logger("Navegando para o formulário de inclusão...")
        
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "content")))
        
        controller_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "CONTROLLER")))
        controller_link.click()

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "content")))
        
        menu_veiculo = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Veículo")))
        actions = ActionChains(driver)
        actions.move_to_element(menu_veiculo).perform() 
        
        menu_compras = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Compras sem cartão")))
        menu_compras.click()
        
        menu_inclusao = wait.until(EC.visibility_of_element_located((By.LINK_TEXT, "Inclusão")))
        menu_inclusao.click()
        
        logger("Navegação do menu concluída.")
        return True
    except Exception as e:
        logger(f"ERRO CRÍTICO ao navegar para o formulário: {e}")
        return False

# --- Função Principal de Preenchimento ---

def preencher_um_registro(driver, wait, dados_combinados, logger):
    try:
        wait.until(EC.presence_of_element_located(LOCATORS['cliente_codigo']))
        
        logger("Preenchendo campos fixos (Cliente/Contrato)...")
        driver.find_element(*LOCATORS['cliente_codigo']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['cliente_codigo']).send_keys("3359")
        driver.find_element(*LOCATORS['contrato_informe']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['contrato_informe']).send_keys("00101033590125")

        logger("Preenchendo dados do motorista...")
        driver.find_element(*LOCATORS['motorista_nome']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['motorista_nome']).send_keys(dados_combinados['nome'])
        driver.find_element(*LOCATORS['motorista_matricula']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['motorista_matricula']).send_keys(dados_combinados['matricula'])
        
        # --- Lógica do Popup (chamando sua função) ---
        placa = dados_combinados['placa']
        _buscar_placa_popup(driver, placa, logger) 
        
        driver.find_element(*LOCATORS['destino']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['destino']).send_keys(dados_combinados['destino'])

        # --- Dados da Transação (Principal - Diesel) ---
        logger("Preenchendo dados da transação (data, hora, hodômetro)...")
        
        driver.find_element(*LOCATORS['data']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['data']).send_keys(dados_combinados['data'])
        driver.find_element(*LOCATORS['hora']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['hora']).send_keys(dados_combinados['hora_para_preencher'])
        driver.find_element(*LOCATORS['hodometro_abastecimento']).send_keys(Keys.CONTROL, "a")
        driver.find_element(*LOCATORS['hodometro_abastecimento']).send_keys(dados_combinados['hodometro_abastecimento'])

        logger("Preenchendo dados do produto principal (Diesel)...")
        
        campo_valor_diesel = driver.find_elements(*LOCATORS['valor_total'])[0]
        campo_valor_diesel.send_keys(Keys.CONTROL, "a")
        campo_valor_diesel.send_keys(dados_combinados['valor_total'])
        
        campo_litros_diesel = driver.find_elements(*LOCATORS['litros'])[0]
        campo_litros_diesel.send_keys(Keys.CONTROL, "a")
        campo_litros_diesel.send_keys(dados_combinados['litros'])

        elemento_select_diesel = driver.find_elements(*LOCATORS['select_produto'])[0]
        seletor_diesel = Select(elemento_select_diesel)
        seletor_diesel.select_by_value("72") # Padrão Diesel S10
        
        # --- Lógica do Aditivo (se existir) ---
        if 'aditivo' in dados_combinados and dados_combinados['aditivo']:
            aditivo_dados = dados_combinados['aditivo']
            logger(f"Detectado Aditivo/Arla. Preenchendo: {aditivo_dados}")
            
            try:
                driver.find_element(*LOCATORS['btn_incluir_novo']).click()
                
                wait.until(
                    lambda d: len(d.find_elements(*LOCATORS['select_produto'])) > 1,
                    "Timeout: O segundo campo de produto (Arla) não apareceu."
                )
                
                todos_selects_produto = driver.find_elements(*LOCATORS['select_produto'])
                todos_campos_valor = driver.find_elements(*LOCATORS['valor_total'])
                todos_campos_litros = driver.find_elements(*LOCATORS['litros'])

                select_arla_elem = todos_selects_produto[1]
                seletor_arla = Select(select_arla_elem)
                seletor_arla.select_by_value("85") # 85 = ÓLEO ARLA 32
                
                valor_arla_elem = todos_campos_valor[1]
                valor_arla_elem.send_keys(Keys.CONTROL, "a")
                valor_arla_elem.send_keys(aditivo_dados['valor'])
                
                litros_arla_elem = todos_campos_litros[1]
                litros_arla_elem.send_keys(Keys.CONTROL, "a")
                litros_arla_elem.send_keys(aditivo_dados['litros'])
                
                logger("Campos do aditivo preenchidos com sucesso.")
            
            except Exception as e_arla:
                logger(f"ERRO CRÍTICO ao tentar preencher o aditivo/arla: {e_arla}")
                return False
        else:
            logger("Nenhum aditivo detectado para este registro.")
        
        logger("Preenchimento deste registro concluído.")
        return True
        
    except Exception as e:
        logger(f"ERRO ao preencher os campos do formulário: {e}")
        return False
