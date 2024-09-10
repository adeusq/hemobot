# Hemobot - Sistema Automatizado de Processos

## Descrição
O Hemobot é um sistema automatizado desenvolvido em Python para facilitar a manipulação de planilhas Excel, exportação de dados em formato TXT, conversão de arquivos XLS para XLSX e resultados de genotipagem. 

## Funcionalidades

### 1. Automatizar Planilha - Excel
- **Descrição**: Preenche automaticamente uma planilha Excel com dados processados.
- **Uso**: Seleciona o arquivo Excel e o sistema preenche as células com base nos dados fornecidos.

### 2. Exportar Dados de Extração - TXT
- **Descrição**: Exporta os dados de uma planilha Excel para um arquivo TXT.
- **Uso**: Seleciona a planilha Excel e define o nome do arquivo de saída em TXT.

### 3. Converter Arquivo - XLS/XLSX
- **Descrição**: Converte arquivos no formato XLS para XLSX.
- **Uso**: Seleciona o arquivo XLS e o sistema gera o arquivo XLSX correspondente.

### 4. Resultados - Genotipagem
- **Descrição**: Concatena e organiza dados de genotipagem em um único arquivo.
- **Uso**: Seleciona os arquivos de genotipagem para concatenar os dados.

## Interface Gráfica (GUI)
O Hemobot possui uma interface gráfica desenvolvida com Tkinter, onde o usuário pode interagir com as funcionalidades mencionadas através de botões intuitivos.

## Requisitos
- Python 3.7+
- Bibliotecas: `tkinter`, `openpyxl`, `pyautogui`, `pyperclip`, `Pillow`, `pandas`

## Instalação
1. Clone o repositório:

   ```bash
   git clone https://github.com/usuario/hemobot.git

2. Download das dependências:
   ```bash
   pip install -r requirements.txt
