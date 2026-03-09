from db import inicializar_banco
from funcionario import abrir_tela_cadastro
from escalas import (
    abrir_tela_primeira_entrada,
    abrir_dialogo_listar_escalas_global,
    abrir_gerador_com_escala
)
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import sys
import os
from tkinter import Tk, PhotoImage

def resource_path(relative_path):
    base = getattr(sys, '_MEIPASS', os.path.abspath('.'))
    return os.path.join(base, relative_path)

# criação da janela
root = Tk()

# definir ícone com fallback seguro
try:
    root.iconbitmap(resource_path('asset/APP_logo.ico'))
except Exception:
    try:
        img = PhotoImage(file=resource_path('asset/logo.png'))
        root.iconphoto(False, img)
    except Exception:
        pass



def abrir_cadastro():
    abrir_tela_cadastro()

def main():
    root = tk.Tk()
    root.title("Sistema de Escalas - Menu Principal")
    root.geometry("800x600")

    root.configure(bg="#FAFAFA")

    root.iconbitmap("asset/APP_logo.ico")

    root.columnconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    style = ttk.Style()
    style.configure("Custom.TFrame", background="#FAFBF9")

    frame = ttk.Frame(root, style="Custom.TFrame", padding=20)
    frame.grid(row=0, column=0, sticky="nsew")

    frame.columnconfigure(0, weight=1)



    titulo = ttk.Label(frame, text="Menu Principal", font=("Arial", 25),anchor="center")
    titulo.grid(row=0, column=0, pady=20,sticky="ew")

    img = Image.open(r"asset\APP_logo.png")

    img = img.resize((150, 150), Image.Resampling.LANCZOS)

    logo = ImageTk.PhotoImage(img)

    lbl_logo = ttk.Label(frame, image=logo)
    lbl_logo.grid(row=1, column=0, pady=10)

    style = ttk.Style()
    style.configure(
        "Custom.TButton",
        font=("Arial", 16, "bold"),  # fonte maior
        padding=10,  # aumenta o tamanho
        foreground="#0072FF",  # cor do texto
        background="white"  # cor de fundo (pode variar conforme tema)
    )

    btn_cadastro = ttk.Button(frame, text="Cadastro de Funcionários", style="Custom.TButton",command=abrir_cadastro)
    btn_cadastro.grid(row=2, column=0, pady=10, sticky="ew")
    btn_escalas = ttk.Button(frame, text="Gerar escalas", style="Custom.TButton", command=abrir_tela_primeira_entrada)
    btn_escalas.grid(row=3, column=0, pady=10, sticky="ew")
    # botão para consultar escalas salvas
    btn_consultar = ttk.Button(
        frame,
        text="Consultar Escalas",
        style="Custom.TButton",
        command=lambda: abrir_dialogo_listar_escalas_global(root, abrir_gerador_com_escala)
    )
    btn_consultar.grid(row=4, column=0, pady=10, sticky="ew")


    root.mainloop()

if __name__ == "__main__":
    main()


