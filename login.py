import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import os
from PIL import Image, ImageTk
from gui import mostrar_menu_principal

LOGIN_FILE = "last_login.txt"

SESSION_TIMEOUT_MINUTES = 30  # Tempo de expiração da sessão em minutos

# Mapeamento de usuários e senhas
USUARIOS_SENHAS = {
    "hemobot": "H3m0b0t@2024!"    
}

def verificar_login(usuario, senha):
    return USUARIOS_SENHAS.get(usuario) == senha

def preencher_senha(usuario_entry, senha_entry):
    usuario = usuario_entry.get()
    senha = USUARIOS_SENHAS.get(usuario)
    if senha:
        senha_entry.delete(0, tk.END)  
        senha_entry.insert(0, senha)  
        senha_entry.config(show="*")   

def salvar_data_login():
    with open(LOGIN_FILE, "w") as f:
        f.write(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

def carregar_data_login():
    if not os.path.exists(LOGIN_FILE):
        return None
    with open(LOGIN_FILE, "r") as f:
        return f.read().strip()

def login():
    def checar_credenciais():
        usuario = usuario_entry.get()
        senha = senha_entry.get()
        if verificar_login(usuario, senha):
            messagebox.showinfo("Bem-vindo!", "Login efetuado com sucesso.\nClique em 'OK' para acessar o painel.")
            salvar_data_login()
            root.destroy()
            mostrar_menu_principal()
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos. Tente novamente.")

    def set_placeholder(entry, placeholder):
        entry.insert(0, placeholder)
        entry.config(fg='grey')
        entry.bind("<FocusIn>", lambda e: on_focus_in(e, placeholder))
        entry.bind("<FocusOut>", lambda e: on_focus_out(e, placeholder))

    def on_focus_in(event, placeholder):
        if event.widget.get() == placeholder:
            event.widget.delete(0, tk.END)
            event.widget.config(fg='black')

    def on_focus_out(event, placeholder):
        if event.widget.get() == "":
            set_placeholder(event.widget, placeholder)

    def on_usuario_entry_change(event):
        preencher_senha(usuario_entry, senha_entry)

    root = tk.Tk()
    root.title("Hemobot")
    root.iconbitmap('C:/project/hemobot/icons8-bot-16.ico')

    largura, altura = 400, 350
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - largura) // 2
    y = (screen_height - altura) // 2
    root.geometry(f"{largura}x{altura}+{x}+{y}")

    root.resizable(False, False)

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(expand=True, fill=tk.BOTH)

    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=3)

    logo = Image.open('C:/project/hemobot/logo-hemobot.png')  
    logo = logo.resize((150, 150), Image.Resampling.LANCZOS)  
    logo_photo = ImageTk.PhotoImage(logo)

    logo_label = tk.Label(frame, image=logo_photo)
    logo_label.grid(row=0, columnspan=2, pady=(0, 15))
    
    placeholder_usuario = "Usuário"
    placeholder_senha = "Senha"

    tk.Label(frame, text="Usuário:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=(0, 5), sticky=tk.E)
    usuario_entry = tk.Entry(frame, font=("Arial", 12), bd=1, relief="flat")
    usuario_entry.grid(row=1, column=1, padx=5, pady=(5, 5), sticky=tk.W)
    set_placeholder(usuario_entry, placeholder_usuario)
    usuario_entry.bind("<KeyRelease>", on_usuario_entry_change)

    tk.Label(frame, text="Senha:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=(0, 5), sticky=tk.E)
    senha_entry = tk.Entry(frame, font=("Arial", 12), show="*", bd=1, relief="flat")
    senha_entry.grid(row=2, column=1, padx=5, pady=(5, 15), sticky=tk.W)
    set_placeholder(senha_entry, placeholder_senha)

    login_button = tk.Button(frame, text="Entrar", font=("Arial", 12), bg="#4CAF50", fg="white", bd=0, relief="flat", padx=10, pady=5, command=checar_credenciais)
    login_button.grid(row=3, columnspan=2, pady=(10, 0))

    root.mainloop()

def verificar_login_necessario():
    data_hoje = datetime.now()
    data_ultimo_login_str = carregar_data_login()
    if data_ultimo_login_str:
        data_ultimo_login = datetime.strptime(data_ultimo_login_str, "%Y-%m-%d %H:%M:%S")
        if (data_hoje - data_ultimo_login) < timedelta(minutes=SESSION_TIMEOUT_MINUTES):
            mostrar_menu_principal()
        else:
            login()
    else:
        login()

if __name__ == "__main__":
    verificar_login_necessario()
