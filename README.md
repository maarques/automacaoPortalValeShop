# 🤖 Automação de Registro de Veículos - Portal Valeshop
Um robô (RPA) desenvolvido em Python que automatiza o processo de registro de múltiplos abastecimentos de veículos no portal Valeshop. O robô lê os dados de uma planilha Excel e preenche o formulário web, lidando com popups, campos dinâmicos e validações de JavaScript.

## ✨ Funcionalidades Principais

Interface Gráfica (GUI): Uma interface simples em Tkinter para seleção de arquivo e log de progresso em tempo real.

Leitura de Planilhas: Utiliza o Pandas para extrair dados do cabeçalho (Condutor, Placa) e uma lista de transações (Data, Valor, Litros) de um arquivo .xlsx.

Automação Web (RPA): Utiliza o Selenium para navegar pelo portal, preencher formulários e submeter os dados.

Lógica de Múltiplos Registros: Faz o loop de todas as transações da planilha, submetendo um formulário para cada uma.

Manipulação de Modal (Popup): Lida com a complexa busca de placas em um modal JavaScript, incluindo IDs duplicados.

Lógica de Produtos Divididos: Interpreta transações que contêm múltiplos produtos (ex: "Diesel + Arla") e preenche os campos dinâmicos de sub-produto.

Configuração Segura: Utiliza um arquivo .env para armazenar a URL de login de forma segura.

## ⚙️ Como Funciona

O fluxo da automação é projetado para ser robusto e lidar com as particularidades do portal Valeshop:

```Bash

1. O usuário seleciona uma planilha Excel (.xlsx) pela interface.
2. O Pandas lê a planilha:
   - Extrai os dados do cabeçalho (Condutor, Placa, etc.).
   - Extrai a tabela de transações (Data, Valor, Litros, etc.).
   - Processa os dados (ex: divide "Diesel + Arla", calcula horas).
3. O Selenium abre o navegador e pausa, aguardando o login manual (com CAPTCHA).
4. O robô navega pelos menus até o formulário de "Inclusão".
5. PARA CADA transação na lista:
   - Preenche os campos principais.
   - Chama a função de busca no modal da Placa (lidando com IDs duplicados).
   - Se a transação tiver Aditivo/Arla, clica em "Incluir novo" e preenche os campos [1].
   - Submete o formulário principal.
   - Aguarda a página limpar para o próximo registro.
6. Ao final, fecha o navegador.
```
## 🚀 Instalação e Configuração
Siga os passos abaixo para configurar e executar o projeto em sua máquina local.

1. Pré-requisitos
Python 3.8+

Google Chrome (ou o navegador correspondente ao seu webdriver)

2. Clonar o Repositório
```Bash

git clone https://github.com/SEU-USUARIO/automacaoportalvaleshop.git
cd automacaoportalvaleshop
```
3. Configurar Ambiente Virtual
É altamente recomendado usar um ambiente virtual (.venv) para isolar as dependências do projeto.

```Bash

# Cria um ambiente virtual
python -m venv .venv

# Ativa o ambiente (Windows)
.\.venv\Scripts\activate

# Ativa o ambiente (Linux/Mac)
# source .venv/bin/activate
```
4. Instalar Dependências
Com o ambiente ativado, instale todas as bibliotecas necessárias listadas no requirements.txt.

```Bash

# Instala todas as bibliotecas
pip install -r requirements.txt
```
## 🛠️ Configuração
Antes de rodar, é necessário configurar a URL de login.

Como a URL de Login faz parte do portal da empresa, eu deixei em um arquivo env e não posso publicar aqui

▶️ Executando o Projeto
Com o ambiente ativado (.venv) e o arquivo .env configurado, basta executar o main.py:

```Bash

python main.py
```
Isso abrirá a interface gráfica. Selecione a planilha Excel e clique em "REGISTRAR VEÍCULO" para iniciar a automação.

## 📁 Estrutura do Projeto
O projeto foi modularizado para separar responsabilidades, tornando a manutenção mais simples:

```Bash

automacaoportalvaleshop/
├── 📂 automation/           # Contém toda a lógica de automação
│   ├── controller.py       # O "cérebro": orquestra o fluxo (login, loop, submit)
│   └── core_functions.py   # O "arquivo de funções": Funções puras de Pandas e Selenium
├── 📂 classes/               # Contém a interface gráfica
│   └── app_gui.py          # A tela principal (Tkinter) e seus callbacks
├── main.py                   # Ponto de entrada: inicializa a GUI
├── .env                      # Arquivo de configuração (URL)
├── .gitignore                # Ignora arquivos desnecessários
├── requirements.txt          # Lista de dependências (pip)
└── README.md
```
