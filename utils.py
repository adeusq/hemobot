import openpyxl
import pyperclip
import pyautogui
from tkinter import Tk, messagebox, filedialog
import tkinter as tk
import time
import pandas as pd
import re
import os

def preencher_planilha(linha_inicio):
    # Inicializa a janela do Tkinter e oculta a janela principal
    root = Tk()
    root.withdraw()
    
    # Abre uma janela de diálogo para o usuário escolher o arquivo Excel
    arquivo_caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    
    # Verifica se o usuário selecionou um arquivo
    if not arquivo_caminho:
        messagebox.showwarning("Nenhum arquivo selecionado", "Por favor, selecione um arquivo para continuar.")
        return
    
    # Carrega a planilha escolhida
    workbook = openpyxl.load_workbook(arquivo_caminho)
    pagina_genotipagem = workbook['doc_automatizado']
    indice_coluna_destino = 0
    indice_coluna_abo = 1
    indice_coluna_rhd = 2
    indice_coluna_fenotipagem = 3
    
    linhas_para_atualizar = [
        linha for linha in pagina_genotipagem.iter_rows(min_row=linha_inicio)
        if linha[indice_coluna_destino].value is None
    ]

    if not linhas_para_atualizar:
        messagebox.showinfo("Nenhuma atualização", "Não há linhas para atualizar a partir da linha especificada.")
        return

    for linha in pagina_genotipagem.iter_rows(min_row=linha_inicio):
        if linha[indice_coluna_destino].value is not None:
            continue

        num_amostra = linha[4].value
        pyperclip.copy(num_amostra)
        
        # Captura da PF
        pyautogui.click(195, 325)
        pyautogui.write('=')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.write('006')
        pyautogui.press('enter')
        pyautogui.sleep(2)
        pyautogui.click(200, 349)        
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')

        info_campo = pyperclip.paste()
        linha[indice_coluna_destino].value = info_campo 
                
        # Captura Tipagem ABO
        pyautogui.click(356, 353)  
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        abo = pyperclip.paste()
        linha[indice_coluna_abo].value = abo

        # Captura Tipagem RhD
        pyautogui.click(449, 350)  
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        rhd = pyperclip.paste()
        linha[indice_coluna_rhd].value = rhd
        
    # Segunda fase: Captura da fenotipagem, com navegação
    for linha in linhas_para_atualizar:
        pf = linha[indice_coluna_destino].value
        if pf is None:
            continue
        
        pyperclip.copy(pf)
        
        # Navegação para a página de fenotipagem
        pyautogui.click(475, 155)  # Clicar na área para acessar a página de fenotipagem
        pyautogui.sleep(1)
        pyautogui.click(481, 341)  # Clicar no campo de entrada de PF
        pyautogui.sleep(1)
        pyautogui.click(697, 342)  # Clicar no campo de entrada de PF
        pyautogui.sleep(1)
        pyautogui.click(650, 272)  # Clicar no campo de entrada de PF
        pyautogui.sleep(1)
        pyautogui.hotkey('ctrl', 'v')  # Colar o número da amostra (PF)
        pyautogui.sleep(2)
        pyautogui.click(641, 485)  # Confirmar busca da fenotipagem
        pyautogui.sleep(1)
        pyautogui.click(201, 362)  # Selecionar informação para cópia
        pyautogui.sleep(1)
        pyautogui.click(243, 526)  # Selecionar informação para cópia
        pyautogui.sleep(1)
        pyautogui.click(192, 329)  # Selecionar informação para cópia
        pyautogui.sleep(5)

        # Captura da informação de fenotipagem
        pyautogui.scroll(-500)  # Rolagem para baixo
        pyautogui.sleep(2)

        pyautogui.click(476, 461)  # Selecionar informação para cópia
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        fenotipagem = pyperclip.paste()
        
        
        # Encontrar a parte após "Fenotipagem" e pegar a próxima linha
        if "Fenotipagem" in fenotipagem:
            partes = fenotipagem.split("Fenotipagem", 1)  # Divide a string em duas partes, separadas por "Fenotipagem"
            linha_fenotipagem = partes[1].strip().split('\n')[0]  # Remove espaços extras e pega a primeira linha após "Fenotipagem"
        else:
            linha_fenotipagem = "Fenotipagem não encontrada"
            
        linha[indice_coluna_fenotipagem].value = linha_fenotipagem

    # Abre janela de diálogo para salvar o arquivo atualizado
    salvar_caminho = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )

    # Salva o arquivo atualizado
    if salvar_caminho:
        workbook.save(salvar_caminho)
        messagebox.showinfo("Arquivo salvo", f"Arquivo salvo com sucesso em {salvar_caminho}.")
    else:
        messagebox.showwarning("Salvamento cancelado", "Arquivo não foi salvo.")     

def preencher_fenotipagem(linha_inicio):
    workbook = openpyxl.load_workbook('todos CORE.xlsx')
    pagina_fenotipagem = workbook['consulta_PF']
    indice_coluna_amostra = 3
    
    linhas_para_atualizar = [
    linha for linha in pagina_fenotipagem.iter_rows(min_row=linha_inicio)
    if linha[indice_coluna_amostra].value is None
    ]
    
    if not linhas_para_atualizar:
        messagebox.showinfo("Nenhuma atualização", "Não há linhas para atualizar a partir da linha especificada.")
        return

    for linha in pagina_fenotipagem.iter_rows(min_row=linha_inicio):
        if linha[indice_coluna_amostra].value is not None:
            continue

        # Captura o número da amostra
        num_pf = linha[0].value
        pyperclip.copy(num_pf)
        pyautogui.click(676, 276)  # Coordenadas onde a amostra é inserida
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(678, 480)  # Coordenadas onde a amostra é inserida
        pyautogui.click(220, 325)  # Coordenadas onde a amostra é inserida
        pyautogui.click(249, 627)  # Coordenadas onde a amostra é inserida
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.sleep(2)
        
        info_campo = pyperclip.paste()

        linha[indice_coluna_amostra].value = info_campo 
        
    workbook.save('todos CORE.xlsx')  # Salve o arquivo após as atualizações
    messagebox.showinfo("Sucesso", "A fenotipagem foi atualizada com sucesso!")
    
def atualizar_progresso(progresso, valor):
    progresso['value'] = valor
    progresso.update_idletasks()

def export_columns_to_txt(consulta_file, origem_file, consulta_sheet, origem_sheet, txt_file, update_progress=None):
    codigo_column_index_origem = 8  # Índice da coluna do código de extração no arquivo de origem
    amostra_column_index_origem = 4  # Índice da coluna onde está o número da amostra no arquivo de origem

    # Carrega os arquivos de origem e consulta
    workbook_origem = openpyxl.load_workbook(origem_file)
    sheet_origem = workbook_origem[origem_sheet]

    workbook_consulta = openpyxl.load_workbook(consulta_file)
    sheet_consulta = workbook_consulta[consulta_sheet]

    # Cria um conjunto de números de amostra do arquivo de consulta
    amostras_consulta = {str(row[1].value).strip() for row in sheet_consulta.iter_rows(min_row=2) if row[1].value is not None}
    
    # Inicializa o dicionário de códigos de extração
    codigo_para_amostra = {}
    for row in sheet_origem.iter_rows(min_row=2):
        amostra = str(row[amostra_column_index_origem].value).strip()
        codigo = str(row[codigo_column_index_origem].value).strip()

        # Se o número da amostra do arquivo de origem também estiver no arquivo de consulta, adiciona ao dicionário
        if amostra in amostras_consulta:
            codigo_para_amostra[amostra] = codigo

    # Verifica se algum dado foi encontrado
    if not codigo_para_amostra:
        messagebox.showwarning("Aviso", "Nenhum dado correspondente encontrado entre os arquivos.")
        return

    total_linhas = len(codigo_para_amostra)
    
    # Gera o arquivo TXT com base na comparação de números de amostra
    with open(txt_file, 'w', encoding='utf-8') as file:
        for i, (amostra, codigo) in enumerate(codigo_para_amostra.items()):
            file.write(f"{amostra}\t{codigo}\n")
            if update_progress:
                progress = (i + 1) / total_linhas * 100
                update_progress(progress)
                time.sleep(0.05)  # Ajuste conforme necessário para a responsividade

    messagebox.showinfo("Sucesso", f'Dados exportados para {txt_file} com sucesso!')

    
def clean_column_name(col_name):
    return re.sub(r'\s*\(.*?\)\s*', '', col_name)

def processar_fenotipagem(file_path):
    df = pd.read_excel(file_path, sheet_name='ID CORE XT Fenótipo')
    
    resultados = []
    
    for index, row in df.iterrows():
        amostra = row.iloc[0]
        match = re.search(r'B315\d+|B3121\d+', amostra)
        if match:
            amostra_id = match.group()
        else:
            amostra_id = amostra
            
        antigenos = []
        for col in df.columns[1:]:
            value = row[col]
            if value in ['+', '0', 'NC', 'UN'] or isinstance(value, str) and re.match(r'\+\(\d+\)', value):
                col_name = clean_column_name(col)
                antigenos.append(f"{col_name}({value})")

        categories = [
            (0, 9), (9, 15), (15, 17), (17, 19), (19, 25),
            (25, 27), (27, 31), (31, 33), (33, 35), (35, None)
        ]

        antigenos_str = '; '.join([
            ', '.join(antigenos[start:end] if end is not None else antigenos[start:])
            for start, end in categories
        ])

        resultado = f"{amostra_id}: Fenotipagem deduzida a partir da genotipagem; {antigenos_str}".rstrip('; ').rstrip('.')
        resultados.append(resultado)
        
    # Abrir janela para o usuário escolher onde salvar o arquivo
    output_file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Arquivo de Texto", "*.txt")],
        title="Salvar arquivo de resultados"
    )

    if output_file_path:
            try:
                with open(output_file_path, 'w', encoding='utf-8') as f:
                    for resultado in resultados:
                        f.write(resultado + '\n\n')
                    messagebox.showinfo("Sucesso", f"Dados concatenados com sucesso e salvos em: {output_file_path}")
            except Exception as e:
                    messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo: {e}")
    else:
        messagebox.showwarning("Cancelado", "Operação de salvamento cancelada.")
            
    return output_file_path 

def converter_xls_para_xlsx(xls_file_path):
    # Configura a janela principal do Tkinter
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    # Abre uma caixa de diálogo para o usuário escolher onde salvar o arquivo e o nome do arquivo
    xlsx_file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar como"
    )

    if not xlsx_file_path:
        messagebox.showwarning("Cancelado", "Operação de salvamento cancelada.")
        return None

    try:
        # Abre o arquivo Excel original
        xls = pd.ExcelFile(xls_file_path)
        
        with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
            for sheet_name in xls.sheet_names:
                # Lê a planilha
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                if sheet_name == 'ID CORE XT Fenótipo':
                    # Remove as linhas de 1 a 19 (considerando que o índice é baseado em zero, isso exclui as linhas 2 a 20)
                    df = df.drop(index=range(0, 19))
                    
                    # Mantém apenas as linhas até a linha 68 (ou seja, mantém as primeiras 68 linhas após a remoção)
                    df = df.head(50)
                    
                    # Remove a linha de números (se estiver presente)
                    df = df.reset_index(drop=True)
                
                # Escreve o DataFrame no novo arquivo Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em: {xlsx_file_path}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo: {e}")

    return xlsx_file_path