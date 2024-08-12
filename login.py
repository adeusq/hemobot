import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import os
from PIL import Image, ImageTk
from gui import mostrar_menu_principal

LOGIN_FILE = "last_login.txt"

def verificar_login(usuario, senha):
    # Aqui você pode verificar o login contra um banco de dados ou arquivo
    return usuario == "lab" and senha == "geno2024"

def salvar_data_login():
    with open(LOGIN_FILE, "w") as f:
        f.write(datetime.now().strftime("%Y-%m-%d"))

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
            messagebox.showinfo("Login", "Login bem-sucedido!")
            salvar_data_login()
            root.destroy()
            mostrar_menu_principal()
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos!")
            
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
            set_placeholder(event.widget, placeholder)          # type: ignore

    root = tk.Tk()
    root.title("Hemobot")
    root.iconbitmap('C:/project/hemobot/icons8-bot-16.ico')

    # Define o tamanho da janela e centraliza
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

    # Carregar e configurar a imagem da logo
    logo = Image.open('C:/project/hemobot/logo-hemobot.png')  
    logo = logo.resize((150, 150), Image.Resampling.LANCZOS)  # Redimensiona a logo
    logo_photo = ImageTk.PhotoImage(logo)

    # Adicionar a imagem da logo
    logo_label = tk.Label(frame, image=logo_photo)
    logo_label.grid(row=0, columnspan=2, pady=(0, 15))
    
    placeholder_usuario = "Usuário"
    placeholder_senha = "Senha"

    tk.Label(frame, text="Usuário:", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=(0, 5), sticky=tk.E)
    usuario_entry = tk.Entry(frame, font=("Arial", 12), bd=1, relief="flat")
    usuario_entry.grid(row=1, column=1, padx=5, pady=(5, 5), sticky=tk.W)
    set_placeholder(usuario_entry, placeholder_usuario)

    tk.Label(frame, text="Senha:", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=(0, 5), sticky=tk.E)
    senha_entry = tk.Entry(frame, font=("Arial", 12), show="*", bd=1, relief="flat")
    senha_entry.grid(row=2, column=1, padx=5, pady=(5, 15), sticky=tk.W)
    set_placeholder(senha_entry, placeholder_senha)

    login_button = tk.Button(frame, text="Entrar", font=("Arial", 12), bg="#4CAF50", fg="white", bd=0, relief="flat", padx=10, pady=5, command=checar_credenciais)
    login_button.grid(row=3, columnspan=2, pady=(10, 0))

    root.mainloop()

def verificar_login_necessario():
    data_hoje = datetime.now().strftime("%Y-%m-%d")
    data_ultimo_login = carregar_data_login()
    if data_ultimo_login != data_hoje:
        login()
    else:
        mostrar_menu_principal()

if __name__ == "__main__":
    verificar_login_necessario()
