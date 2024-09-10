import openpyxl
import pyperclip
import pyautogui
from tkinter import Tk, messagebox, filedialog
import tkinter as tk
import time
import pandas as pd
import re

def preencher_planilha(linha_inicio):
    root = Tk()
    root.withdraw()
    
    arquivo_caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    
    if not arquivo_caminho:
        messagebox.showwarning("Nenhum arquivo selecionado", "Por favor, selecione um arquivo para continuar.")
        return
    
    workbook = openpyxl.load_workbook(arquivo_caminho)
    pagina_genotipagem = workbook['Extraction-Log']
    indice_coluna_destino = 0
    indice_coluna_abo = 2
    indice_coluna_rhd = 3
    indice_coluna_fenotipagem = 4
    
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

        num_amostra = linha[1].value
        pyperclip.copy(num_amostra)
        
        # Captura da PF
        pyautogui.click(243, 399)
        pyautogui.write('=')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.write('006')
        pyautogui.press('enter')
        pyautogui.sleep(2)
        pyautogui.click(251, 426)        
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')

        info_campo = pyperclip.paste()
        linha[indice_coluna_destino].value = info_campo 
                
        # Captura Tipagem ABO
        pyautogui.click(440, 429)  
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        abo = pyperclip.paste()
        linha[indice_coluna_abo].value = abo

        # Captura Tipagem RhD
        pyautogui.click(555, 428)  
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        rhd = pyperclip.paste()
        linha[indice_coluna_rhd].value = rhd
        
    salvar_caminho = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )

    if salvar_caminho:
        workbook.save(salvar_caminho)
        messagebox.showinfo("Arquivo salvo", f"Arquivo salvo com sucesso em {salvar_caminho}.")
    else:
        messagebox.showwarning("Salvamento cancelado", "Arquivo não foi salvo.")     
    
def atualizar_progresso(progresso, valor):
    progresso['value'] = valor
    progresso.update_idletasks()

def export_columns_to_txt(txt_file_input, origem_file, origem_sheet, txt_file_output, update_progress=None):
    codigo_column_index_origem = 8  # Índice da coluna do código de extração no arquivo de origem
    amostra_column_index_origem = 1  # Índice da coluna onde está o número da amostra no arquivo de origem

    workbook_origem = openpyxl.load_workbook(origem_file)
    sheet_origem = workbook_origem[origem_sheet]

    codigo_para_amostra = {}
    for row in sheet_origem.iter_rows(min_row=2):
        codigo = str(row[codigo_column_index_origem].value).strip()
        amostra = str(row[amostra_column_index_origem].value).strip()

        if codigo and amostra:
            codigo_para_amostra[codigo] = amostra

    if not codigo_para_amostra:
        messagebox.showwarning("Aviso", "Nenhum dado correspondente encontrado no arquivo de origem.")
        return

    with open(txt_file_input, 'r', encoding='utf-8') as infile, open(txt_file_output, 'w', encoding='utf-8') as outfile:
        linhas = infile.readlines()
        total_linhas = len(linhas)

        for i, linha in enumerate(linhas):
            codigo = linha.strip()
            amostra = codigo_para_amostra.get(codigo, "Amostra não encontrada")

            outfile.write(f"{amostra}\t{codigo}\n")

            if update_progress:
                progress = (i + 1) / total_linhas * 100
                update_progress(progress)
                time.sleep(0.05)  
                
    messagebox.showinfo("Sucesso", f'Dados exportados para {txt_file_output} com sucesso!')
   
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
    root = tk.Tk()
    root.withdraw() 

    xlsx_file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")],
        title="Salvar como"
    )

    if not xlsx_file_path:
        messagebox.showwarning("Cancelado", "Operação de salvamento cancelada.")
        return None

    try:
        xls = pd.ExcelFile(xls_file_path)
        
        with pd.ExcelWriter(xlsx_file_path, engine='openpyxl') as writer:
            for sheet_name in xls.sheet_names:
                
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                if sheet_name == 'ID CORE XT Fenótipo':
                    
                    df = df.drop(index=range(0, 19))
                    
                    df = df.head(50)
                    
                    df = df.reset_index(drop=True)
                
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em: {xlsx_file_path}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o arquivo: {e}")

    return xlsx_file_path