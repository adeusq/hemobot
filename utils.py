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
        linha[indice_coluna_fenotipagem].value = fenotipagem
    
    salvar_caminho = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    
    if not salvar_caminho:
        messagebox.showwarning("Nenhum local de salvamento selecionado", "Por favor, selecione um local para salvar o arquivo.")
        return
    
    # Salva as alterações no arquivo escolhido pelo usuário
    workbook.save(salvar_caminho)
    messagebox.showinfo("Sucesso", "A planilha foi atualizada e salva com sucesso!")

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

def export_columns_to_txt(excel_file, sheet_name, txt_file, update_progress=None):
    codigo_column_index = 8  
    amostra_column_index = 1  
    
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook[sheet_name]

    with open(txt_file, 'r', encoding='utf-8') as file:
        linhas_txt = file.readlines()

    codigo_para_amostra = {}
    for row in sheet.iter_rows(min_row=2): 
        codigo = row[codigo_column_index].value
        amostra = row[amostra_column_index].value
        if codigo is not None and amostra is not None:
            codigo_para_amostra[str(codigo)] = str(amostra)

    total_linhas = len(linhas_txt)
    with open(txt_file, 'w', encoding='utf-8') as file:
        for i, linha in enumerate(linhas_txt):
            codigo_txt = linha.strip()
            if codigo_txt in codigo_para_amostra:
                amostra = codigo_para_amostra[codigo_txt]
                file.write(f"{amostra}\t{codigo_txt}\n")
            else:
                file.write(f"{codigo_txt}\n")
            if update_progress:
                progress = (i + 1) / total_linhas * 100
                update_progress(progress)
                time.sleep(0.05)

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

    output_file_path = 'resultados_fenotipagem.txt'
    with open(output_file_path, 'w', encoding='utf-8') as f:
        for resultado in resultados:
            f.write(resultado + '\n\n')

    return output_file_path  

def converter_xls_para_xlsx(xls_file_path):
    xls = pd.ExcelFile(xls_file_path)
    xlsx_file_path = f'{os.path.splitext(xls_file_path)[0]}.xlsx'
    with pd.ExcelWriter(xlsx_file_path) as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return xlsx_file_path
