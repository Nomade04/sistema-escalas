import tkinter as tk
from tkinter import ttk, messagebox
import mysql.connector
import unicodedata
import re

def conectar():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="sistema_escalas"
    )

def normalizar_texto(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    return s.upper()

def cadastrar_funcionario(nome, cpf, setor, turno):
    conn = conectar()
    cursor = conn.cursor()
    setor_norm = normalizar_texto(setor)
    turno_norm = normalizar_texto(turno)
    sql = "INSERT INTO funcionarios (nome, cpf, setor, turno) VALUES (%s, %s, %s, %s)"
    valores = (nome, cpf, setor_norm, turno_norm)
    try:
        cursor.execute(sql, valores)
        conn.commit()
        messagebox.showinfo("Sucesso", f"Funcionário {nome} cadastrado com sucesso!")
    except mysql.connector.Error as err:
        messagebox.showerror("Erro", f"Erro ao cadastrar funcionário: {err}")
    finally:
        cursor.close()
        conn.close()

def editar_funcionario(tree):
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um funcionário para editar.")
        return

    valores = tree.item(selecionado[0], "values")
    funcionario_id = valores[0]
    nome_atual = valores[1]
    cpf_atual = valores[2]
    setor_atual = valores[3]
    turno_atual = valores[4]

    def on_salvar_edicao():
        novo_nome = entry_nome.get().strip()
        novo_cpf = entry_cpf.get().strip()
        novo_setor = combo_setor.get().strip()
        novo_turno = combo_turno.get().strip()

        if not novo_nome or not novo_cpf or not novo_setor or not novo_turno:
            messagebox.showwarning("Aviso", "Preencha todos os campos.")
            return

        conn = conectar()
        cursor = conn.cursor()
        try:
            sql = "UPDATE funcionarios SET nome = %s, cpf = %s, setor = %s, turno = %s WHERE id = %s"
            valores = (novo_nome, novo_cpf, normalizar_texto(novo_setor), normalizar_texto(novo_turno), funcionario_id)
            cursor.execute(sql, valores)
            conn.commit()
            messagebox.showinfo("Sucesso", f"Funcionário {novo_nome} atualizado com sucesso!")
            listar_funcionarios(tree)
            win_editar.destroy()
        except mysql.connector.Error as err:
            messagebox.showerror("Erro", f"Erro ao atualizar funcionário: {err}")
        finally:
            cursor.close()
            conn.close()

    win_editar = tk.Toplevel()
    win_editar.title("Editar Funcionário")
    win_editar.geometry("360x220")
    frm = ttk.Frame(win_editar, padding=10)
    frm.grid(row=0, column=0, sticky="nsew")

    ttk.Label(frm, text="Nome:").grid(row=0, column=0, sticky="w")
    entry_nome = ttk.Entry(frm, width=30)
    entry_nome.grid(row=0, column=1, pady=4)
    entry_nome.insert(0, nome_atual)

    ttk.Label(frm, text="CPF:").grid(row=1, column=0, sticky="w")
    entry_cpf = ttk.Entry(frm, width=20)
    entry_cpf.grid(row=1, column=1, pady=4)
    entry_cpf.insert(0, cpf_atual)

    ttk.Label(frm, text="Setor:").grid(row=2, column=0, sticky="w")
    combo_setor = ttk.Combobox(frm, values=["Padaria","Açougue","Liderança","Loja"], width=20)
    combo_setor.grid(row=2, column=1, pady=4)
    combo_setor.set(setor_atual)

    ttk.Label(frm, text="Turno:").grid(row=3, column=0, sticky="w")
    combo_turno = ttk.Combobox(frm, values=["Manhã","Tarde"], width=20)
    combo_turno.grid(row=3, column=1, pady=4)
    combo_turno.set(turno_atual)

    ttk.Button(frm, text="Salvar", command=on_salvar_edicao).grid(row=4, column=0, columnspan=2, pady=10)

def deletar_funcionario(tree):
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um funcionário para deletar.")
        return

    valores = tree.item(selecionado[0], "values")
    funcionario_id = valores[0]
    nome = valores[1]

    confirmar = messagebox.askyesno("Confirmação", f"Deseja realmente deletar o funcionário {nome}?")
    if not confirmar:
        return

    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM funcionarios WHERE id = %s", (funcionario_id,))
        conn.commit()
        messagebox.showinfo("Sucesso", f"Funcionário {nome} deletado com sucesso!")
        listar_funcionarios(tree)
    except mysql.connector.Error as err:
        messagebox.showerror("Erro", f"Erro ao deletar funcionário: {err}")
    finally:
        cursor.close()
        conn.close()

def listar_funcionarios(tree):
    conn = conectar()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, nome, cpf, setor, turno FROM funcionarios ORDER BY nome")
    resultados = cursor.fetchall()
    conn.close()

    for item in tree.get_children():
        tree.delete(item)

    for f in resultados:
        tree.insert("", "end", values=(f["id"], f["nome"], f["cpf"], f["setor"], f["turno"]))

# Interface Tkinter
def abrir_tela_cadastro():
    janela = tk.Toplevel()  # cria nova janela a partir do menu
    janela.title("Cadastro de Funcionários")

    # se tiver ícone, mantém; caso contrário, ignora erro
    try:
        janela.iconbitmap("asset/APP_logo.ico")
    except Exception:
        pass

    frame = ttk.Frame(janela, padding=10)
    frame.grid(row=0, column=0, sticky="nsew")

    ttk.Label(frame, text="Nome:").grid(row=0, column=0, sticky="w")
    entry_nome = ttk.Entry(frame, width=30)
    entry_nome.grid(row=0, column=1)

    ttk.Label(frame, text="CPF:").grid(row=1, column=0, sticky="w")
    entry_cpf = ttk.Entry(frame, width=20)
    entry_cpf.grid(row=1, column=1)

    ttk.Label(frame, text="Setor:").grid(row=2, column=0, sticky="w")
    combo_setor = ttk.Combobox(frame, values=["Padaria","Açougue","Liderança","Loja"], width=20)
    combo_setor.grid(row=2, column=1)

    ttk.Label(frame, text="Turno:").grid(row=3, column=0, sticky="w")
    combo_turno = ttk.Combobox(frame, values=["Manhã","Tarde"], width=20)
    combo_turno.grid(row=3, column=1)

    def on_cadastrar():
        nome = entry_nome.get()
        cpf = entry_cpf.get()
        setor = combo_setor.get()
        turno = combo_turno.get()
        if not nome or not cpf or not setor or not turno:
            messagebox.showwarning("Aviso", "Preencha todos os campos.")
            return
        cadastrar_funcionario(nome, cpf, setor, turno)
        listar_funcionarios(tree)

    ttk.Button(frame, text="Cadastrar", command=on_cadastrar).grid(row=4, column=0, columnspan=2, pady=10)

    cols = ("ID","Nome","CPF","Setor","Turno")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=120 if col != "Nome" else 220)
    tree.grid(row=5, column=0, columnspan=2, pady=(4,8))

    ttk.Button(frame, text="Listar Funcionários", command=lambda: listar_funcionarios(tree)).grid(row=6, column=0, pady=6, sticky="ew")
    ttk.Button(frame, text="Editar Funcionário", command=lambda: editar_funcionario(tree)).grid(row=6, column=1, pady=6, sticky="ew")
    ttk.Button(frame, text="Deletar Funcionário", command=lambda: deletar_funcionario(tree)).grid(row=7, column=0, columnspan=2, pady=6, sticky="ew")