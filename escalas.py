import calendar
import unicodedata
import re
import mysql.connector
import holidays
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from tkinter.scrolledtext import ScrolledText

import json
import os
import tempfile
import datetime

# PDF generation
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.units import mm
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

# -----------------------
# Constantes e utilitários
# -----------------------

# Meses em Português (índice 1..12)
MESES_PT = [
    None,
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

def normalizar_texto(s: str) -> str:
    if s is None:
        return ""
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^A-Za-z0-9 ]", "", s)
    return s.upper()

MAPA_VARIACOES_SETOR = {
    "ACOUGUE": "ACOUGUE",
    "ACOU GUE": "ACOUGUE",
    "LIDERANCA": "LIDERANCA",
    "LIDER": "LIDERANCA",
    "GERENCIA": "LIDERANCA",
    "PADARIA": "PADARIA",
    "DONO":"LIDERANCA",
    "LOJA": "LOJA",
    "VENDA": "LOJA",
    "VENDAS": "LOJA",
}

MAPA_VARIACOES_TURNO = {
    "MANHA": "MANHA",
    "MANHAO": "MANHA",
    "MANHÃ": "MANHA",
    "M": "MANHA",
    "TARDE": "TARDE",
    "T": "TARDE",
}

SETORES_OBRIGATORIOS = {"PADARIA", "ACOUGUE", "LIDERANCA", "LOJA"}
TURNOS_VALIDOS = {"MANHA", "TARDE"}

def participants_safe(participantes):
    """Retorna lista segura de participantes (evita None)."""
    return participantes or []

# -----------------------
# Conexão com o banco
# -----------------------

def conectar():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="root",
        database="sistema_escalas"
    )

# -----------------------
# Parsing / DB helpers
# -----------------------

def parse_folgas_aggregadas(folgas_text):
    """
    "1:(1,5,8),2:(3,7,10)" -> {1:[1,5,8], 2:[3,7,10]}
    Retorna {} em caso de string vazia ou inválida.
    """
    if not folgas_text:
        return {}
    escala = {}
    s = folgas_text.strip()
    pattern = re.compile(r'(\d+)\s*:\s*\(\s*([0-9,\s]*)\s*\)')
    for m in pattern.finditer(s):
        fid = int(m.group(1))
        dias_str = m.group(2).strip()
        if dias_str == "":
            dias = []
        else:
            dias = [int(x) for x in dias_str.split(',') if x.strip().isdigit()]
        escala[fid] = sorted(set(dias))
    return escala

def extrair_ids_de_folgas(folgas_text):
    """
    Extrai os ids de funcionários presentes na string agregada de folgas.
    Ex: "1:(1,5,8),2:(3,7,10)" -> [1,2]
    """
    if not folgas_text:
        return []
    ids = []
    for m in re.finditer(r'(\d+)\s*:\s*\(', folgas_text):
        try:
            ids.append(int(m.group(1)))
        except Exception:
            pass
    return ids

def listar_escalas_no_banco():
    """
    Retorna lista de tuplas (id, ano, mes, folgas) ordenadas por ano/mes desc.
    """
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, ano, mes, folgas FROM escalas ORDER BY ano DESC, mes DESC")
        rows = cursor.fetchall()
        cursor.close()
        conn.close()
        return rows
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao listar escalas: {e}")
        return []

def obter_escala_por_id(reg_id):
    """
    Retorna dict {fid: [dias]} para o registro com id = reg_id.
    """
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT folgas FROM escalas WHERE id = %s", (reg_id,))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if not row:
            return {}
        folgas_text = row[0] or ""
        return parse_folgas_aggregadas(folgas_text)
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao obter escala: {e}")
        return {}

def deletar_escala_por_id(reg_id):
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM escalas WHERE id = %s", (reg_id,))
        conn.commit()
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        messagebox.showerror("Erro", f"Falha ao deletar escala: {e}")
        return False

def obter_funcionarios_por_ids(ids):
    if not ids:
        return []
    conn = conectar()
    cursor = conn.cursor(dictionary=True)
    formato = ",".join(["%s"] * len(ids))
    sql = f"SELECT id, nome, setor, turno FROM funcionarios WHERE id IN ({formato})"
    cursor.execute(sql, ids)
    resultados = cursor.fetchall()
    conn.close()
    return resultados

# -----------------------
# Validações e geração inicial
# -----------------------

def normalizar_setor_turno(funcionario):
    raw_setor = funcionario.get("setor", "")
    raw_turno = funcionario.get("turno", "")
    norm_setor = normalizar_texto(raw_setor)
    norm_turno = normalizar_texto(raw_turno)
    setor_final = MAPA_VARIACOES_SETOR.get(norm_setor, norm_setor)
    turno_final = MAPA_VARIACOES_TURNO.get(norm_turno, norm_turno)
    return setor_final, turno_final

def validar_setores_por_turno(funcionarios):
    presentes_por_turno = {}  # turno -> set(setores)
    for f in funcionarios:
        setor, turno = normalizar_setor_turno(f)
        if not turno or turno not in TURNOS_VALIDOS:
            turno = "SEM_TURNO"
        presentes_por_turno.setdefault(turno, set()).add(setor)

    insuficiencia_por_turno = {}
    for turno in ("MANHA", "TARDE"):
        setores_presentes = presentes_por_turno.get(turno, set())
        faltantes = SETORES_OBRIGATORIOS - setores_presentes
        if faltantes:
            insuficiencia_por_turno[turno] = faltantes

    if "SEM_TURNO" in presentes_por_turno:
        insuficiencia_por_turno["SEM_TURNO"] = set()

    setores_totais = set()
    for s in presentes_por_turno.values():
        setores_totais.update(s)
    ok_total = SETORES_OBRIGATORIOS.issubset(setores_totais)

    return ok_total, insuficiencia_por_turno, presentes_por_turno

def obter_dias_e_feriados(mes, ano, estado=None):
    dias_no_mes = calendar.monthrange(ano, mes)[1]
    br_holidays = holidays.Brazil(years=ano, prov=estado) if estado else holidays.Brazil(years=ano)
    feriados_dias = [f"{data.day:02d}" for data in br_holidays if data.month == mes]
    feriados_str = ",".join(sorted(set(feriados_dias), key=lambda x: int(x)))
    return dias_no_mes, feriados_str

def primeira_entrada_por_ids(mes, ano, estado, ids_participantes):
    funcionarios = obter_funcionarios_por_ids(ids_participantes)

    if len(funcionarios) != len(ids_participantes):
        encontrados = {f['id'] for f in funcionarios}
        faltantes_ids = [str(i) for i in ids_participantes if i not in encontrados]
        raise ValueError(f"Alguns IDs não foram encontrados no banco: {', '.join(faltantes_ids)}")

    setores_presentes = { normalizar_setor_turno(f)[0] for f in funcionarios }
    faltantes_globais = SETORES_OBRIGATORIOS - setores_presentes
    if faltantes_globais:
        faltantes_legivel = ", ".join(sorted([s.capitalize() for s in faltantes_globais]))
        raise ValueError(f"É necessário pelo menos um funcionário de cada setor (Padaria, Açougue, Liderança e Loja). Faltam: {faltantes_legivel}")

    ok_total, insuf_por_turno, presentes_por_turno = validar_setores_por_turno(funcionarios)

    dias_no_mes, feriados_str = obter_dias_e_feriados(mes, ano, estado)

    avisos = []
    for turno, faltantes in insuf_por_turno.items():
        if turno == "SEM_TURNO":
            avisos.append("Existem participantes sem turno definido; corrija para MANHA ou TARDE.")
            continue
        if faltantes:
            faltantes_legivel = ", ".join(sorted([s.capitalize() for s in faltantes]))
            avisos.append(f"Turno {turno.capitalize()}: faltam setores: {faltantes_legivel}")

    return {
        "mes": mes,
        "ano": ano,
        "estado": estado,
        "dias_no_mes": dias_no_mes,
        "feriados": feriados_str,
        "participantes": funcionarios,
        "presenca_por_turno": {t: sorted(list(s)) for t, s in presentes_por_turno.items()},
        "insuficiencia_por_turno": {t: sorted(list(s)) for t, s in insuf_por_turno.items()},
        "avisos": avisos,
        "ok_total_setores": ok_total
    }

# -----------------------
# Salvamento agregado (uma linha por mês)
# -----------------------

def salvar_escala_no_banco_agregado(escala, ano, mes, participantes,
                                    dias_no_mes=None, feriados_text=None,
                                    tabela="escalas"):
    """
    Grava uma única linha por mês em `tabela` com coluna `folgas` contendo:
    "1:(1,5,8),2:(3,7,10),..."
    """
    if escala is None or not isinstance(escala, dict):
        messagebox.showerror("Erro", "Formato da escala inválido.")
        return False

    try:
        ano_i = int(ano)
        mes_i = int(mes)
    except Exception:
        messagebox.showerror("Erro", "Ano ou mês inválido.")
        return False

    if dias_no_mes is None:
        try:
            dias_no_mes = calendar.monthrange(ano_i, mes_i)[1]
        except Exception:
            dias_no_mes = 0

    ultimo_dia = int(dias_no_mes) if dias_no_mes else calendar.monthrange(ano_i, mes_i)[1]
    erros = []

    escala_norm = {}
    for k, v in escala.items():
        try:
            fid = int(k)
        except Exception:
            erros.append(f"ID inválido na escala: {k}")
            continue
        dias = []
        for d in (v or []):
            try:
                di = int(d)
                dias.append(di)
            except Exception:
                erros.append(f"Funcionário {fid}: dia inválido '{d}'")
        dias = sorted(set(dias))
        for d in dias:
            if d < 1 or d > ultimo_dia:
                erros.append(f"Funcionário {fid}: dia {d} fora do mês (1..{ultimo_dia}).")
        escala_norm[fid] = [d for d in dias if 1 <= d <= ultimo_dia]

    participantes_ids = {int(p['id']) for p in participantes}
    for pid in participantes_ids:
        if not escala_norm.get(pid):
            nome = next((p['nome'] for p in participants_safe(participantes) if int(p['id']) == pid), str(pid))
            erros.append(f"Funcionário {pid} - {nome} não possui folga marcada.")

    for pid in participantes_ids:
        dias = escala_norm.get(pid, [])
        pontos = [0] + dias + [ultimo_dia + 1]
        for i in range(1, len(pontos)):
            intervalo = pontos[i] - pontos[i-1] - 1
            if intervalo > 6:
                nome = next((p['nome'] for p in participants_safe(participantes) if int(p['id']) == pid), str(pid))
                erros.append(f"Funcionário {pid} - {nome}: intervalo de {intervalo} dias sem folga (maior que 6).")

    if erros:
        texto = "Foram encontrados os seguintes problemas:\n\n" + "\n".join(f"- {e}" for e in erros)
        messagebox.showerror("Erros na validação da escala", texto)
        return False

    partes = []
    for fid in sorted(escala_norm.keys()):
        dias = escala_norm.get(fid, [])
        parte = f"{fid}:({','.join(str(d) for d in sorted(dias))})"
        partes.append(parte)
    folgas_aggregadas = ",".join(partes)

    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {tabela} WHERE ano = %s AND mes = %s", (ano_i, mes_i))
        cursor.execute(
            f"INSERT INTO {tabela} (mes, ano, dias_no_mes, feriados, folgas) VALUES (%s, %s, %s, %s, %s)",
            (mes_i, ano_i, dias_no_mes, feriados_text or "", folgas_aggregadas)
        )
        conn.commit()
        cursor.close()
        conn.close()
        messagebox.showinfo("Sucesso", "Escala salva no banco com sucesso!")
        return True

    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        messagebox.showerror("Erro ao salvar no banco", f"Ocorreu um erro ao salvar a escala:\n{e}")
        return False

# -----------------------
# UI: Primeira entrada (seleção de participantes)
# -----------------------

def abrir_tela_primeira_entrada():
    def carregar_funcionarios():
        try:
            conn = conectar()
            cur = conn.cursor(dictionary=True)
            cur.execute("SELECT id, nome, setor, turno FROM funcionarios ORDER BY nome")
            rows = cur.fetchall()
            conn.close()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar funcionários: {e}")
            return

        listbox_func.delete(0, tk.END)
        for f in rows:
            listbox_func.insert(tk.END, f"{f['id']} - {f['nome']} - {f.get('setor','')} - {f.get('turno','')}")

    def obter_ids_selecionados():
        sel = listbox_func.curselection()
        if not sel:
            return []
        ids = []
        for i in sel:
            texto = listbox_func.get(i)
            id_str = texto.split(" - ", 1)[0]
            try:
                ids.append(int(id_str))
            except ValueError:
                pass
        return ids

    def on_gerar():
        try:
            mes = int(entry_mes.get())
            ano = int(entry_ano.get())
            if not (1 <= mes <= 12):
                raise ValueError("Mês inválido")
        except Exception:
            messagebox.showwarning("Aviso", "Informe mês (1-12) e ano válidos.")
            return

        estado = entry_estado.get().strip() or None

        sel = listbox_func.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione ao menos um funcionário.")
            return
        ids = []
        for i in sel:
            texto = listbox_func.get(i)
            id_str = texto.split(" - ", 1)[0]
            try:
                ids.append(int(id_str))
            except ValueError:
                pass
        if not ids:
            messagebox.showwarning("Aviso", "Seleção inválida. Recarregue a lista e tente novamente.")
            return

        try:
            resultado = primeira_entrada_por_ids(mes, ano, estado, ids)
        except ValueError as ve:
            messagebox.showerror("Validação", str(ve))
            return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar primeira entrada: {e}")
            return

        win.resultado = resultado
        btn_gerar_escala.config(state="normal")

        txt_result.delete("1.0", tk.END)
        txt_result.insert(tk.END, f"Mês: {resultado['mes']:02d}/{resultado['ano']}\n")
        txt_result.insert(tk.END, f"Dias no mês: {resultado['dias_no_mes']}\n")
        txt_result.insert(tk.END, f"Feriados (dias): {resultado['feriados'] or 'Nenhum'}\n\n")

        txt_result.insert(tk.END, "Participantes selecionados:\n")
        for p in resultado['participantes']:
            txt_result.insert(tk.END, f"  {p['id']} - {p['nome']} - {p.get('setor', '')} - {p.get('turno', '')}\n")
        txt_result.insert(tk.END, "\n")

        txt_result.insert(tk.END, "Presença por turno:\n")
        for t, setores in resultado['presenca_por_turno'].items():
            txt_result.insert(tk.END, f"  {t}: {', '.join(setores) if setores else 'Nenhum'}\n")
        txt_result.insert(tk.END, "\n")

        if resultado['insuficiencia_por_turno']:
            txt_result.insert(tk.END, "Avisos de insuficiência por turno:\n")
            for t, falt in resultado['insuficiencia_por_turno'].items():
                if t == "SEM_TURNO":
                    txt_result.insert(tk.END, "  Existem participantes sem turno definido; corrija-os.\n")
                elif falt:
                    txt_result.insert(tk.END, f"  Turno {t}: faltam setores: {', '.join(falt)}\n")
            txt_result.insert(tk.END, "\n")

        if resultado['avisos']:
            txt_result.insert(tk.END, "Avisos gerais:\n")
            for a in resultado['avisos']:
                txt_result.insert(tk.END, f"  - {a}\n")
            txt_result.insert(tk.END, "\n")

        txt_result.insert(tk.END,
                          f"Validação global de setores completa: {'SIM' if resultado['ok_total_setores'] else 'NÃO'}\n")

    win = tk.Toplevel()
    win.title("Primeira Entrada da Escala")
    win.geometry("800x600")

    left = ttk.Frame(win, padding=8)
    left.grid(row=0, column=0, sticky="nsew")
    win.columnconfigure(0, weight=1)
    win.columnconfigure(1, weight=2)
    win.rowconfigure(0, weight=1)

    ttk.Label(left, text="Funcionários (selecione múltiplos):").grid(row=0, column=0, sticky="w")
    listbox_func = tk.Listbox(left, selectmode=tk.EXTENDED, width=40, height=20)
    listbox_func.grid(row=1, column=0, sticky="nsew", pady=6)
    left.rowconfigure(1, weight=1)

    btn_refresh = ttk.Button(left, text="Carregar funcionários", command=carregar_funcionarios)
    btn_refresh.grid(row=2, column=0, pady=6, sticky="ew")

    right = ttk.Frame(win, padding=8)
    right.grid(row=0, column=1, sticky="nsew")

    frm_params = ttk.Frame(right)
    frm_params.grid(row=0, column=0, sticky="ew")
    frm_params.columnconfigure(1, weight=1)

    ttk.Label(frm_params, text="Mês:").grid(row=0, column=0, sticky="w")
    entry_mes = ttk.Entry(frm_params, width=6)
    entry_mes.grid(row=0, column=1, sticky="w", padx=(4, 10))
    entry_mes.insert(0, str(datetime.datetime.now().month))

    ttk.Label(frm_params, text="Ano:").grid(row=0, column=2, sticky="w")
    entry_ano = ttk.Entry(frm_params, width=8)
    entry_ano.grid(row=0, column=3, sticky="w", padx=(4, 10))
    entry_ano.insert(0, str(datetime.datetime.now().year))

    ttk.Label(frm_params, text="Estado (UF) opcional:").grid(row=1, column=0, sticky="w", pady=(8,0))
    entry_estado = ttk.Entry(frm_params, width=10)
    entry_estado.grid(row=1, column=1, sticky="w", pady=(8,0))

    btn_gerar = ttk.Button(frm_params, text="Gerar Primeira Entrada", command=on_gerar)
    btn_gerar.grid(row=2, column=0, columnspan=4, pady=10, sticky="ew")

    btn_gerar_escala = ttk.Button(
        frm_params,
        text="Gerar Escala",
        state="disabled",
        command=lambda: abrir_tela_gerar_escala(getattr(win, "resultado", None))
    )
    btn_gerar_escala.grid(row=3, column=0, columnspan=4, pady=6, sticky="ew")

    ttk.Label(right, text="Resultado / Avisos:").grid(row=1, column=0, sticky="w", pady=(6,0))
    txt_result = ScrolledText(right, width=60, height=25)
    txt_result.grid(row=2, column=0, sticky="nsew")
    right.rowconfigure(2, weight=1)

    carregar_funcionarios()

# -----------------------
# UI: Gerar / Editar Escala (grade)
# -----------------------

def abrir_tela_gerar_escala(resultado):
    if resultado is None:
        messagebox.showerror("Erro", "Resultado da primeira entrada não encontrado.")
        return

    mes = resultado["mes"]
    ano = resultado["ano"]
    dias_no_mes = resultado["dias_no_mes"]
    participantes = resultado.get("participantes", [])
    estado = resultado.get("estado", None)

    br_holidays = holidays.Brazil(years=ano, prov=estado) if estado else holidays.Brazil(years=ano)
    feriados_map = {}
    for data, nome in br_holidays.items():
        if data.month == mes:
            feriados_map[data.day] = nome

    dias_abrev = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    cor_domingo = "#d9d9d9"
    cor_feriado = "#0b5ea8"
    cor_feriado_text = "white"
    cor_normal = "white"
    cor_dsr = "#bfe9ff"
    cor_dsr_text = "black"

    win = tk.Toplevel()
    nome_mes_pt = MESES_PT[mes] if 1 <= mes <= 12 else calendar.month_name[mes]
    win.title(f"Escala {nome_mes_pt} {ano}")
    win.geometry("1100x650")

    container = ttk.Frame(win, padding=6)
    container.grid(row=0, column=0, sticky="nsew")
    win.rowconfigure(0, weight=1)
    win.columnconfigure(0, weight=1)

    header = ttk.Label(container, text=f"Escala - {nome_mes_pt} {ano}",
                       font=("Arial", 14, "bold"))
    header.grid(row=0, column=0, sticky="w", pady=(0, 8))

    canvas = tk.Canvas(container)
    vsb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    hsb = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
    grid_frame = ttk.Frame(canvas)

    grid_frame_id = canvas.create_window((0, 0), window=grid_frame, anchor="nw")
    canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    canvas.grid(row=1, column=0, sticky="nsew")
    vsb.grid(row=1, column=1, sticky="ns")
    hsb.grid(row=2, column=0, sticky="ew")
    container.rowconfigure(1, weight=1)
    container.columnconfigure(0, weight=1)

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    grid_frame.bind("<Configure>", on_frame_configure)

    tk.Label(grid_frame, text="Funcionário / Dia", borderwidth=1, relief="solid", anchor="center", width=36).grid(row=0,
                                                                                                                  column=0,
                                                                                                                  sticky="nsew")

    for d in range(1, dias_no_mes + 1):
        weekday = calendar.weekday(ano, mes, d)
        nome_abrev = dias_abrev[weekday]
        if d in feriados_map:
            lbl = tk.Label(grid_frame, text=f"{d}\n{nome_abrev}", borderwidth=1, relief="solid", width=8, height=2,
                           bg=cor_feriado, fg=cor_feriado_text)
        elif weekday == 6:
            lbl = tk.Label(grid_frame, text=f"{d}\n{nome_abrev}", borderwidth=1, relief="solid", width=8, height=2,
                           bg=cor_domingo)
        else:
            lbl = tk.Label(grid_frame, text=f"{d}\n{nome_abrev}", borderwidth=1, relief="solid", width=8, height=2)
        lbl.grid(row=0, column=d, sticky="nsew")

    celulas = {}

    def toggle_celula(func_id, dia):
        btn = celulas[(func_id, dia)]
        estado_atual = getattr(btn, "_estado", "empty")
        if estado_atual != "dsr":
            try:
                btn._orig_bg = btn.cget("bg")
            except Exception:
                btn._orig_bg = cor_normal
            try:
                btn._orig_fg = btn.cget("fg")
            except Exception:
                btn._orig_fg = "black"
            try:
                btn._orig_text = btn.cget("text")
            except Exception:
                btn._orig_text = ""
            btn.configure(bg=cor_dsr, fg=cor_dsr_text, text="DSR")
            btn._estado = "dsr"
        else:
            orig_bg = getattr(btn, "_orig_bg", cor_normal)
            orig_fg = getattr(btn, "_orig_fg", "black")
            orig_text = getattr(btn, "_orig_text", "")
            btn.configure(bg=orig_bg, fg=orig_fg, text=orig_text)
            if dia in feriados_map:
                btn._estado = "feriado"
            else:
                weekday = calendar.weekday(ano, mes, dia)
                if weekday == 6:
                    btn._estado = "domingo"
                else:
                    btn._estado = "empty"

    for r, f in enumerate(participantes, start=1):
        func_id = f.get("id")
        nome = f.get("nome", "")
        setor_text = f.get("setor", "") or ""
        turno_text = f.get("turno", "") or ""
        lbl_text = f"{func_id} - {nome}"
        if setor_text:
            lbl_text += f" - {setor_text}"
        if turno_text:
            lbl_text += f" - {turno_text}"
        lbl_nome = tk.Label(grid_frame, text=lbl_text, borderwidth=1, relief="solid", anchor="w", width=36)
        lbl_nome.grid(row=r, column=0, sticky="nsew")

        for d in range(1, dias_no_mes + 1):
            weekday = calendar.weekday(ano, mes, d)
            btn = tk.Button(grid_frame, text="", width=8, height=2, relief="raised")
            if d in feriados_map:
                btn.configure(bg=cor_feriado, fg=cor_feriado_text, text="FER")
                btn._estado = "feriado"
            elif weekday == 6:
                btn.configure(bg=cor_domingo, fg="black", text="")
                btn._estado = "domingo"
            else:
                btn.configure(bg=cor_normal, fg="black", text="")
                btn._estado = "empty"

            btn.configure(command=lambda fid=func_id, day=d: toggle_celula(fid, day))
            btn.grid(row=r, column=d, sticky="nsew", padx=0, pady=0)
            celulas[(func_id, d)] = btn

    win.update_idletasks()

    folgas_brutas = resultado.get("_folgas_brutas")
    if folgas_brutas:
        escala_dict = parse_folgas_aggregadas(folgas_brutas)
        aplicar_escala_na_grade(celulas, escala_dict)

    for c in range(0, dias_no_mes + 1):
        grid_frame.grid_columnconfigure(c, weight=1)

    rodape = ttk.Frame(container, padding=(0, 8, 0, 0))
    rodape.grid(row=3, column=0, sticky="ew")
    rodape.columnconfigure(0, weight=1)

    if feriados_map:
        txt_feriados = tk.Text(rodape, height=3, wrap="word", bg=win.cget("bg"), borderwidth=0)
        linhas = []
        for dia in sorted(feriados_map.keys()):
            nome = feriados_map[dia]
            linhas.append(f"{dia:02d} — {nome}")
        txt_feriados.insert("1.0", "Feriados do mês: " + "  |  ".join(linhas))
        txt_feriados.configure(state="disabled")
        txt_feriados.grid(row=0, column=0, sticky="ew", pady=(6, 4))
    else:
        ttk.Label(rodape, text="Nenhum feriado no mês.").grid(row=0, column=0, sticky="w", pady=(6, 4))

    def coletar_escala():
        escala = {}
        try:
            for (fid, dia), btn in celulas.items():
                try:
                    fid_i = int(fid)
                    dia_i = int(dia)
                except Exception:
                    continue
                if getattr(btn, "_estado", "") == "dsr":
                    escala.setdefault(fid_i, []).append(dia_i)
        except Exception:
            return {}

        for fid in list(escala.keys()):
            dias = sorted(set(int(d) for d in escala[fid]))
            escala[fid] = dias
        return escala

    def salvar_escala_no_banco(escala, ano, mes, participantes):
        return salvar_escala_no_banco_agregado(
            escala, ano, mes, participantes,
            dias_no_mes=dias_no_mes, feriados_text=resultado.get('feriados', "")
        )

    def on_salvar():
        escala = coletar_escala()
        if escala is None:
            messagebox.showerror("Erro", "Falha ao coletar a escala.")
            return
        if not escala:
            messagebox.showwarning("Aviso", "Nenhuma folga marcada.")
            return

        ok = salvar_escala_no_banco_agregado(
            escala,
            resultado['ano'],
            resultado['mes'],
            resultado.get('participantes', []),
            dias_no_mes=resultado.get('dias_no_mes'),
            feriados_text=resultado.get('feriados', ""),
            tabela="escalas"
        )
        if ok:
            pass

    btn_frame = ttk.Frame(rodape)
    btn_frame.grid(row=1, column=0, sticky="e", pady=(4, 0))
    ttk.Button(btn_frame, text="Salvar Escala (resumo)", command=on_salvar).grid(row=0, column=0, padx=4)
    ttk.Button(btn_frame, text="Fechar", command=win.destroy).grid(row=0, column=1, padx=4)

    win.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

# -----------------------
# Funções auxiliares de UI/grade
# -----------------------

def limpar_marcas_da_grade(celulas):
    """
    Restaura todas as células ao estado original salvo em atributos _orig_*
    """
    for (fid, dia), btn in list(celulas.items()):
        if hasattr(btn, "_orig_bg"):
            try:
                btn.configure(bg=btn._orig_bg)
            except Exception:
                pass
        if hasattr(btn, "_orig_text"):
            try:
                btn.configure(text=btn._orig_text)
            except Exception:
                pass
        btn._estado = getattr(btn, "_orig_estado", "")

def aplicar_escala_na_grade(celulas, escala_dict, cor_dsr="#bfe9ff", texto_dsr="DSR"):
    """
    Marca as células correspondentes como 'dsr' e atualiza visual.
    celulas: dict {(fid, dia): btn}
    escala_dict: {fid: [dias]}
    """
    limpar_marcas_da_grade(celulas)
    for fid, dias in escala_dict.items():
        for d in dias:
            key = (int(fid), int(d))
            btn = celulas.get(key)
            if not btn:
                continue
            if not hasattr(btn, "_orig_saved"):
                btn._orig_saved = True
                try:
                    btn._orig_bg = btn.cget("bg")
                except Exception:
                    btn._orig_bg = "white"
                try:
                    btn._orig_text = btn.cget("text")
                except Exception:
                    btn._orig_text = ""
                btn._orig_estado = getattr(btn, "_estado", "")
            btn._estado = "dsr"
            try:
                btn.configure(bg=cor_dsr, text=texto_dsr)
            except Exception:
                pass

# -----------------------
# Exportar PDF: função principal (pede local ao usuário ou salva em Desktop/pdf)
# -----------------------

def gerar_pdf_para_escala(reg_id, dias_por_pagina=15, salvar_em=None, parent_window=None):
    """
    Gera um PDF da escala com blocos de dias de tamanho `dias_por_pagina`.
    reg_id: id da tabela 'escalas' (registro agregado por mês)
    dias_por_pagina: quantos dias mostrar por página/segmento (int)
    salvar_em: caminho do arquivo de saída (opcional). Se None, salva em ~/Desktop/pdf (cria se necessário).
    parent_window: janela pai para diálogos (opcional)
    Retorna caminho do arquivo gerado ou None em caso de erro.
    """
    if not REPORTLAB_AVAILABLE:
        messagebox.showerror("Erro", "Biblioteca reportlab não está instalada. Instale 'reportlab' para exportar PDF.")
        return None

    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT ano, mes, dias_no_mes, feriados, folgas FROM escalas WHERE id = %s", (reg_id,))
        row = cursor.fetchone()
        cursor.close()
        conn.close()
        if not row:
            messagebox.showerror("Erro", "Registro não encontrado no banco.")
            return None
        ano_r, mes_r, dias_no_mes_r, feriados_r, folgas_r = row
        ano_r = int(ano_r)
        mes_r = int(mes_r)
        dias_no_mes = int(dias_no_mes_r) if dias_no_mes_r else calendar.monthrange(ano_r, mes_r)[1]
        feriados_map = {}
        br_holidays = holidays.Brazil(years=ano_r)
        for data, nome in br_holidays.items():
            if data.month == mes_r:
                feriados_map[data.day] = nome
        folgas_dict = parse_folgas_aggregadas(folgas_r or "")
        ids = extrair_ids_de_folgas(folgas_r or "")
        funcionarios = obter_funcionarios_por_ids(ids) if ids else []
        id_to_func = {f['id']: f for f in funcionarios}

        # se salvar_em não foi informado, salvar em Desktop/pdf (cria pasta se necessário)
        if salvar_em is None:
            home = os.path.expanduser("~")
            desktop_dir = os.path.join(home, "Desktop")
            # Em sistemas em português a pasta pode se chamar "Área de Trabalho", mas Desktop costuma existir.
            # Tentativa alternativa: procurar "Área de Trabalho" se "Desktop" não existir
            if not os.path.isdir(desktop_dir):
                alt_desktop = os.path.join(home, "Área de Trabalho")
                if os.path.isdir(alt_desktop):
                    desktop_dir = alt_desktop
                else:
                    # fallback para home
                    desktop_dir = home
            pdf_dir = os.path.join(desktop_dir, "pdf")
            try:
                os.makedirs(pdf_dir, exist_ok=True)
            except Exception:
                # fallback para temp se não conseguir criar na área de trabalho
                pdf_dir = tempfile.gettempdir()
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            salvar_em = os.path.join(pdf_dir, f"escala_{ano_r}_{mes_r}_{reg_id}_{timestamp}.pdf")

        # preparar documento
        doc = SimpleDocTemplate(salvar_em, pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm,
                                topMargin=10*mm, bottomMargin=10*mm)
        styles = getSampleStyleSheet()
        story = []

        titulo = Paragraph(f"<b>Escala - {MESES_PT[mes_r]} {ano_r}</b>", styles['Title'])
        story.append(titulo)
        story.append(Spacer(1, 4*mm))

        # dividir dias em blocos
        dias = list(range(1, dias_no_mes + 1))
        blocos = [dias[i:i + dias_por_pagina] for i in range(0, len(dias), dias_por_pagina)]

        for bloco_idx, bloco in enumerate(blocos, start=1):
            bloco_title = Paragraph(f"<b>Período {bloco[0]:02d} a {bloco[-1]:02d}</b>", styles['Heading3'])
            story.append(bloco_title)
            story.append(Spacer(1, 2*mm))

            header = ["Funcionário"]
            header.extend([f"{d:02d}" for d in bloco])
            data = [header]

            participantes_ord = [id_to_func[i] for i in ids if i in id_to_func] if ids else sorted(funcionarios, key=lambda x: x.get('nome',''))
            if not participantes_ord and folgas_dict:
                participantes_ord = []
                for fid in sorted(folgas_dict.keys()):
                    participantes_ord.append(id_to_func.get(fid, {"id": fid, "nome": str(fid), "setor": "", "turno": ""}))

            for p in participantes_ord:
                fid = p.get('id')
                nome = p.get('nome', '')
                setor = p.get('setor', '') or ''
                turno = p.get('turno', '') or ''
                nome_exib = f"{fid} - {nome}"
                if setor:
                    nome_exib += f" - {setor}"
                if turno:
                    nome_exib += f" - {turno}"
                row = [nome_exib]
                folgas = folgas_dict.get(fid, [])
                for d in bloco:
                    if d in folgas:
                        row.append("DSR")
                    elif d in feriados_map:
                        row.append("FER")
                    else:
                        row.append("")
                data.append(row)

            table = Table(data, repeatRows=1)
            tbl_style = TableStyle([
                ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
                ('BACKGROUND', (0,0), (-1,0), colors.lightblue),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ])
            for col_idx, d in enumerate(bloco, start=1):
                if d in feriados_map:
                    tbl_style.add('BACKGROUND', (col_idx,0), (col_idx,-1), colors.whitesmoke)
                    tbl_style.add('TEXTCOLOR', (col_idx,0), (col_idx,-1), colors.darkblue)
            table.setStyle(tbl_style)
            story.append(table)
            story.append(Spacer(1, 6*mm))

        gerado_em = Paragraph(f"Gerado em: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal'])
        story.append(Spacer(1, 4*mm))
        story.append(gerado_em)

        doc.build(story)

        messagebox.showinfo("Exportado", f"PDF gerado em:\n{salvar_em}")
        return salvar_em

    except Exception as e:
        messagebox.showerror("Erro ao gerar PDF", f"Ocorreu um erro ao gerar o PDF:\n{e}")
        return None

# -----------------------
# Diálogo: listar / carregar / deletar escalas salvas (com botão Exportar PDF)
# -----------------------

def abrir_dialogo_listar_escalas_global(root, abrir_gerador_callback):
    """
    root: janela principal (Tk)
    abrir_gerador_callback: função que abre a tela de gerar escala e aceita um dict 'resultado_preparado'
    """
    rows = listar_escalas_no_banco()
    if not rows:
        messagebox.showinfo("Escalas", "Nenhuma escala salva encontrada.")
        return

    dlg = tk.Toplevel(root)
    dlg.title("Escalas salvas")
    dlg.transient(root)
    dlg.grab_set()
    dlg.geometry("900x400")

    tree = ttk.Treeview(dlg, columns=("id","ano","mes","folgas"), show="headings")
    tree.heading("id", text="ID")
    tree.heading("ano", text="Ano")
    tree.heading("mes", text="Mês")
    tree.heading("folgas", text="Folgas")
    tree.column("id", width=50, anchor="center")
    tree.column("ano", width=60, anchor="center")
    tree.column("mes", width=80, anchor="center")
    tree.column("folgas", width=600, anchor="w")
    tree.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
    dlg.rowconfigure(0, weight=1)
    dlg.columnconfigure(0, weight=1)

    for r in rows:
        reg_id, ano_r, mes_r, folgas_r = r
        tree.insert("", "end", iid=str(reg_id), values=(reg_id, ano_r, mes_r, folgas_r or ""))

    btn_frame_local = ttk.Frame(dlg)
    btn_frame_local.grid(row=1, column=0, sticky="e", padx=8, pady=8)

    def on_carregar():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma escala para carregar.")
            return
        reg_id = int(sel[0])
        try:
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute("SELECT ano, mes, dias_no_mes, feriados, folgas FROM escalas WHERE id = %s", (reg_id,))
            row = cursor.fetchone()
            cursor.close()
            conn.close()
            if not row:
                messagebox.showerror("Erro", "Registro não encontrado.")
                return
            ano_r, mes_r, dias_no_mes_r, feriados_r, folgas_r = row
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao obter escala: {e}")
            return

        resultado_preparado = {
            "ano": int(ano_r),
            "mes": int(mes_r),
            "dias_no_mes": int(dias_no_mes_r) if dias_no_mes_r else None,
            "feriados": feriados_r or "",
            "participantes": [],
            "_folgas_brutas": folgas_r or ""
        }

        ids = extrair_ids_de_folgas(folgas_r or "")
        if ids:
            try:
                funcionarios = obter_funcionarios_por_ids(ids)
                id_to_func = {f['id']: f for f in funcionarios}
                participantes_list = [id_to_func[i] for i in ids if i in id_to_func]
                resultado_preparado["participantes"] = participantes_list
            except Exception:
                resultado_preparado["participantes"] = []

        abrir_gerador_callback(resultado_preparado)
        dlg.destroy()

    def on_deletar():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma escala para deletar.")
            return
        reg_id = int(sel[0])
        if not messagebox.askyesno("Confirmar", f"Deletar escala id {reg_id}?"):
            return
        ok = deletar_escala_por_id(reg_id)
        if ok:
            tree.delete(str(reg_id))

    def on_exportar_pdf():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma escala para exportar.")
            return
        reg_id = int(sel[0])
        dias_por_pagina = simpledialog.askinteger("Dias por página", "Quantos dias por bloco (ex.: 7, 10, 15)?", parent=dlg, minvalue=1, maxvalue=31)
        if not dias_por_pagina:
            return
        # Salvar automaticamente em Desktop/pdf (cria a pasta se necessário)
        gerar_pdf_para_escala(reg_id, dias_por_pagina, salvar_em=None, parent_window=dlg)

    ttk.Button(btn_frame_local, text="Carregar", command=on_carregar).grid(row=0, column=0, padx=4)
    ttk.Button(btn_frame_local, text="Exportar PDF", command=on_exportar_pdf).grid(row=0, column=1, padx=4)
    ttk.Button(btn_frame_local, text="Deletar", command=on_deletar).grid(row=0, column=2, padx=4)
    ttk.Button(btn_frame_local, text="Fechar", command=dlg.destroy).grid(row=0, column=3, padx=4)

def abrir_gerador_com_escala(resultado_preparado):
    abrir_tela_gerar_escala(resultado_preparado)