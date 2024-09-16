import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading
from utils import converter_xls_para_xlsx, export_columns_to_txt, preencher_planilha, processar_fenotipagem, criar_arquivo_modelo
import os

def centralizar_janela(root, largura, altura):
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - largura) // 2
    y = (screen_height - altura) // 2
    root.geometry(f"{largura}x{altura}+{x}+{y}")

def mostrar_sobre():
    messagebox.showinfo("Hemobot - Sobre", 
        "Hemobot - Sistema de Automação de Processos\n\n"
        "O Hemobot automatiza processos com planilhas em rotina laboratorial de genotipagem.")

def mostrar_ajuda():
    messagebox.showinfo("Hemobot - Ajuda", 
        "Para suporte, consulte a documentação ou entre em contato conosco.")
    
def fechar_sistema(root):
    root.destroy()

def voltar_menu(root):
    root.destroy()
    mostrar_menu_principal()

def mostrar_barra_progresso(txt_file_input, origem_file, txt_file_output):
    def tarefa_longas():
        def update_progress(value):
            progresso['value'] = value
            root.update_idletasks()

        export_columns_to_txt(txt_file_input, origem_file, 'DNA extraídos', txt_file_output, update_progress)
        root.destroy()

    root = tk.Tk()
    root.title("Exportando Dados")
    root.geometry("400x150")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    label = tk.Label(frame, text="Exportando dados para TXT...", font=("Arial", 10))
    label.pack(pady=10)

    progresso = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
    progresso.pack(pady=20)

    thread = threading.Thread(target=tarefa_longas)
    thread.start()

    root.mainloop()

def botao_download():
    root = tk.Tk()
    root.title("Hemobot - Sistema de Automação de Processos")
    botao_modelo = tk.Button(root, text="Baixar Arquivo Modelo para Automação", command=criar_arquivo_modelo)
    botao_modelo.pack(pady=20)
    root.mainloop()
    botao_download()

def exportar_dados_txt():
    txt_file_input = filedialog.askopenfilename(filetypes=[("Arquivos TXT", "*.txt")], title="Selecione o arquivo TXT contendo o código de extração")
    if not txt_file_input:
        messagebox.showwarning("Cancelado", "Operação cancelada.")
        return

    origem_file = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")], title="Selecione o arquivo de busca contendo a aba 'DNA extraídos'")
    if not origem_file:
        messagebox.showwarning("Cancelado", "Operação de origem cancelada.")
        return

    mostrar_barra_progresso(txt_file_input, origem_file, txt_file_input)

def concatenar_dados():
    xls_file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls"), ("Todos os arquivos", "*.*")], title="Selecione o arquivo de resultado")

    if not xls_file_path:
        messagebox.showwarning("Nenhum arquivo selecionado", "Por favor, selecione um arquivo para gerar resultados.")
        return

    xlsx_file_path = converter_xls_para_xlsx(xls_file_path)
    
    if xlsx_file_path:
        output_directory = os.path.dirname(xls_file_path)
        processar_fenotipagem(xlsx_file_path, output_directory)
        
        try:
            os.remove(xlsx_file_path)
        except Exception as e:
            messagebox.showwarning

def mostrar_menu_principal():
    root = tk.Tk()
    root.title("Hemobot - Sistema de Automação de Processos")
    root.geometry("800x600")
    root.resizable(True, True)
    root.iconbitmap('C:/project/hemobot/icons8-bot-16.ico')

    menu_frame = tk.Frame(root)
    menu_frame.pack(side=tk.TOP, fill=tk.X)

    sobre_label = tk.Label(menu_frame, text="Sobre", font=("Arial", 9), cursor="hand2", anchor="center")
    sobre_label.pack(side=tk.LEFT, padx=10)
    sobre_label.bind("<Button-1>", lambda e: mostrar_sobre())

    ajuda_label = tk.Label(menu_frame, text="Ajuda", font=("Arial", 9), cursor="hand2", anchor="center")
    ajuda_label.pack(side=tk.LEFT, padx=10)
    ajuda_label.bind("<Button-1>", lambda e: mostrar_ajuda())

    sair_label = tk.Label(menu_frame, text="Sair", font=("Arial", 9), cursor="hand2")
    sair_label.pack(side=tk.LEFT, padx=10)
    sair_label.bind("<Button-1>", lambda e: fechar_sistema(root))

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    label = tk.Label(frame, text="Olá! Bem-vindo ao Hemobot, um sistema de automação de processos.", font=("Arial", 10))
    label.pack(pady=10)

    option_frame = tk.Frame(frame)
    option_frame.pack(pady=10, fill=tk.X)

    button_style = {'font': ("Arial", 10), 'bg': "#4CAF50", 'fg': "white", 'bd': 0, 'relief': "flat", 'padx': 10, 'pady': 5}

    def acao_baixar():
        criar_arquivo_modelo()
        
    def acao_preencher():
        preencher_planilha(1)
        
    def acao_exportar():
        exportar_dados_txt()

    def acao_concatenar():
        concatenar_dados()

    def criar_opcao(label_text, acao):
        opcao_frame = tk.Frame(option_frame)
        opcao_frame.pack(pady=5, fill=tk.X)

        label = tk.Label(opcao_frame, text=label_text, font=("Arial", 10))
        label.pack(side=tk.LEFT, padx=10)

        iniciar_button = tk.Button(opcao_frame, text="Iniciar", **button_style, command=acao)
        iniciar_button.pack(side=tk.RIGHT, padx=10)

    criar_opcao("Baixar Arquivo Modelo", acao_baixar)
    criar_opcao("Automatizar Planilha - Excel", acao_preencher)
    criar_opcao("Exportar Dados de Extração - TXT", acao_exportar)
    criar_opcao("Resultados - Genotipagem", acao_concatenar)

    centralizar_janela(root, 700, 320)
    root.mainloop()

if __name__ == "__main__":
    mostrar_menu_principal()
