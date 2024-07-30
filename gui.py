import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import threading
from utils import converter_xls_para_xlsx, export_columns_to_txt, preencher_planilha, processar_fenotipagem

def centralizar_janela(root, largura, altura):
    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - largura) // 2
    y = (screen_height - altura) // 2
    root.geometry(f"{largura}x{altura}+{x}+{y}")

def mostrar_sobre():
    messagebox.showinfo("Sobre", "Desenvolvedor: Alba de Deus Moreira\nAno: 2024\nCopyright (c) 2024 Alba de Deus Moreira.\nTodos os direitos reservados.")

def fechar_sistema(root):
    root.destroy()

def voltar_menu(root):
    root.destroy()
    mostrar_menu_principal()

def perguntar_linha_inicio():
    def iniciar_script():
        linha_inicio = linha_inicio_entry.get()
        if linha_inicio.isdigit():
            preencher_planilha(int(linha_inicio))
            voltar_menu(root)
        else:
            messagebox.showerror("Erro", "Por favor, insira um número válido.")

    root = tk.Tk()
    root.title("Preencher Planilha - Hemobot")
    root.resizable(False, False)
    root.iconbitmap('icons8-bot-16.ico')

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    label = tk.Label(frame, text="Digite a partir de qual linha começar:", font=("Arial", 10))
    label.pack(pady=10)

    linha_inicio_entry = tk.Entry(frame, font=("Arial", 10), width=15)
    linha_inicio_entry.pack(pady=5)

    iniciar_button = tk.Button(frame, text="Iniciar", font=("Arial", 10), command=iniciar_script, bg="#4CAF50", fg="white", bd=0, relief="flat", padx=10, pady=5)
    iniciar_button.pack(pady=20)
    
    centralizar_janela(root, 400, 200)
    root.mainloop()

def mostrar_barra_progresso(txt_file):
    def tarefa_longas():
        def update_progress(value):
            progresso['value'] = value
            root.update_idletasks()

        export_columns_to_txt('GENOTIPAGEM - Doadores e pacientes UNIFICADA.xlsx', 'DNA extraídos', txt_file, update_progress)
        root.destroy()

    root = tk.Tk()
    root.title("Exportando Dados")
    root.geometry("400x200")
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

def exportar_dados_txt():
    txt_file = filedialog.asksaveasfilename(filetypes=[("Arquivos TXT", "*.txt")])
    if txt_file:
        mostrar_barra_progresso(txt_file)

def concatenar_dados():
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if file_path:
        output_file_path = processar_fenotipagem(file_path)
        messagebox.showinfo("Sucesso", f'Dados concatenados com sucesso e salvos em: {output_file_path}')

def converter_xls():
    xls_file_path = filedialog.askopenfilename(filetypes=[("Arquivos XLS", "*.xls")])
    if xls_file_path:
        xlsx_file_path = converter_xls_para_xlsx(xls_file_path)
        messagebox.showinfo("Sucesso", f'Arquivo convertido com sucesso e salvo em: {xlsx_file_path}')

def mostrar_menu_principal():
    root = tk.Tk()
    root.title("Hemobot - Sistema de Automação")
    root.resizable(False, False)
    root.iconbitmap('icons8-bot-16.ico')

    # Criar o frame para as opções de menu
    menu_frame = tk.Frame(root)
    menu_frame.pack(side=tk.TOP, fill=tk.X)

    # Adicionar opções de menu no frame
    sobre_label = tk.Label(menu_frame, text="Sobre", font=("Arial", 9), cursor="hand2")
    sobre_label.pack(side=tk.LEFT, padx=10)
    sobre_label.bind("<Button-1>", lambda e: mostrar_sobre())

    ajuda_label = tk.Label(menu_frame, text="Ajuda", font=("Arial", 9), cursor="hand2")
    ajuda_label.pack(side=tk.LEFT, padx=10)
    ajuda_label.bind("<Button-1>", lambda e: messagebox.showinfo("Ajuda", "Bem-vindo à seção de ajuda do Hemobot!\n\nAqui você encontrará informações úteis para navegar e utilizar todas as funcionalidades do nosso sistema de automação do Hemoce.\n\nFuncionalidades:\n\n- Preencher Planilha: Automatize o preenchimento de planilhas Excel.\n\n- Exportar Dados: Exporte dados em formato TXT.\n\n- Converter Arquivos: Converta arquivos de XLS para XLSX.\n\n- Concatenar Dados: Combine dados de genotipagem em um único arquivo.\n\nSuporte:\n\nE-mail: albadedeus@hemobot.com\nTelefone: (85) 99944-3470"))

    sair_label = tk.Label(menu_frame, text="Sair", font=("Arial", 9), cursor="hand2")
    sair_label.pack(side=tk.LEFT, padx=10)
    sair_label.bind("<Button-1>", lambda e: fechar_sistema(root))

    # Criar o frame principal para o conteúdo
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    label = tk.Label(frame, text="Olá! Bem-vindo ao Hemobot, um sistema de automação!", font=("Arial", 10))
    label.pack(pady=10)

    # Criar o frame para as opções com botões "Iniciar"
    option_frame = tk.Frame(frame)
    option_frame.pack(pady=10, fill=tk.X)

    # Adicionar estilo aos botões
    button_style = {'font': ("Arial", 10), 'bg': "#4CAF50", 'fg': "white", 'bd': 0, 'relief': "flat", 'padx': 10, 'pady': 5}

    # Funções para ações
    def acao_preencher():
        perguntar_linha_inicio()

    def acao_exportar():
        exportar_dados_txt()

    def acao_converter():
        converter_xls()

    def acao_concatenar():
        concatenar_dados()

    # Adicionar opções com botões "Iniciar"
    def criar_opcao(label_text, acao):
        opcao_frame = tk.Frame(option_frame)
        opcao_frame.pack(pady=5, fill=tk.X)

        label = tk.Label(opcao_frame, text=label_text, font=("Arial", 10))
        label.pack(side=tk.LEFT, padx=10)

        iniciar_button = tk.Button(opcao_frame, text="Iniciar", **button_style, command=acao)
        iniciar_button.pack(side=tk.RIGHT, padx=10)

    criar_opcao("Automatizar Planilha - Excel", acao_preencher)
    criar_opcao("Exportar Dados - TXT", acao_exportar)
    criar_opcao("Converter XLS para XLSX", acao_converter)
    criar_opcao("Concatenar Dados - Genotipagem", acao_concatenar)

    centralizar_janela(root, 500, 300)
    root.mainloop()

if __name__ == "__main__":
    mostrar_menu_principal()
