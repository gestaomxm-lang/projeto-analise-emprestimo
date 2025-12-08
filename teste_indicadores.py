import os
import json
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

# Pastas e arquivos para persistência dos dados e do relatório
DATA_FOLDER = "dados"
RELATORIOS_FOLDER = "relatorios"
JSON_FILE_PATH = os.path.join(DATA_FOLDER, "registros.json")
REPORT_FILE_PATH = os.path.join(RELATORIOS_FOLDER, "relatorio_cumulativo.txt")

# Cria as pastas se não existirem
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)
if not os.path.exists(RELATORIOS_FOLDER):
    os.makedirs(RELATORIOS_FOLDER)

# Metas fixas para a Unidade HOSPITAL CASA SÃO BERNARDO
META_PACIENTE_CM = 69         # Meta para Paciente dia - CM
META_PACIENTE_CTI = 67        # Meta para Paciente dia - CTI
META_CONSUMO_MATERIAIS = 17056.76    # Meta para Consumo Estoque - Material Hospitalar (R$)
META_CONSUMO_MEDICAMENTOS = 16053.80 # Meta para Consumo Estoque - Medicamentos (R$)

# Lista para armazenar os registros cumulativos (será carregada do arquivo JSON, se existir)
records = []

def load_records():
    """Carrega os registros salvos, se o arquivo JSON existir."""
    global records
    if os.path.exists(JSON_FILE_PATH):
        try:
            with open(JSON_FILE_PATH, "r", encoding="utf-8") as f:
                records = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar os registros:\n{e}")
            records = []

def save_records():
    """Salva a lista de registros no arquivo JSON."""
    try:
        with open(JSON_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=4)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar os registros:\n{e}")

def update_report_file():
    """Atualiza (ou gera) o arquivo de relatório cumulativo em formato texto."""
    content = ""
    if not records:
        content = "Nenhum registro disponível."
    else:
        for rec in records:
            content += f"Data: {rec['Data']}\n"
            content += f"  Paciente dia - CM: {rec['Paciente CM']} -> {rec['Status Paciente CM']}\n"
            content += f"  Paciente dia - CTI: {rec['Paciente CTI']} -> {rec['Status Paciente CTI']}\n"
            content += f"  Consumo Estoque - Material Hospitalar: {format_currency(rec['Consumo Materiais'])} -> {rec['Status Consumo Materiais']}\n"
            content += f"  Consumo Estoque - Medicamentos: {format_currency(rec['Consumo Medicamentos'])} -> {rec['Status Consumo Medicamentos']}\n"
            content += "-" * 50 + "\n"
    try:
        with open(REPORT_FILE_PATH, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar o relatório:\n{e}")

def parse_float(valor_str):
    """
    Converte uma string para float, substituindo a vírgula pelo ponto.
    Exemplo: "15663,20" -> 15663.20
    """
    try:
        return float(valor_str.replace(',', '.'))
    except ValueError:
        raise ValueError("Valor inválido. Use vírgula para separar decimais, ex: 15663,20.")

def format_currency(value):
    """
    Formata um número para o padrão monetário brasileiro.
    Exemplo: 15999.50 -> R$15.999,50
    """
    us_format = format(value, ",.2f")
    return "R$" + us_format.replace(",", "X").replace(".", ",").replace("X", ".")

def submit_data():
    """Coleta os dados da interface, calcula os status dos indicadores e registra o dia, atualizando os arquivos."""
    data = entry_data.get()
    
    try:
        paciente_cm = float(entry_paciente_cm.get().replace(',', '.'))
        paciente_cti = float(entry_paciente_cti.get().replace(',', '.'))
        consumo_materiais = parse_float(entry_consumo_materiais.get())
        consumo_medicamentos = parse_float(entry_consumo_medicamentos.get())
    except ValueError as e:
        messagebox.showerror("Erro", f"Insira valores numéricos válidos para todos os campos.\n{e}")
        return

    # Cálculo do status para os indicadores de pacientes:
    # Se o valor for maior que a meta: "Acima da meta"
    # Se for igual à meta: "Na meta"
    # Se for menor que a meta: "Abaixo da meta"
    if paciente_cm > META_PACIENTE_CM:
        status_paciente_cm = "Acima da meta"
    elif paciente_cm == META_PACIENTE_CM:
        status_paciente_cm = "Na meta"
    else:
        status_paciente_cm = "Abaixo da meta"

    if paciente_cti > META_PACIENTE_CTI:
        status_paciente_cti = "Acima da meta"
    elif paciente_cti == META_PACIENTE_CTI:
        status_paciente_cti = "Na meta"
    else:
        status_paciente_cti = "Abaixo da meta"

    # Cálculo do status para os indicadores de consumo:
    if consumo_materiais < META_CONSUMO_MATERIAIS:
        status_consumo_materiais = "Abaixo da meta"
    elif consumo_materiais == META_CONSUMO_MATERIAIS:
        status_consumo_materiais = "Na meta"
    else:
        status_consumo_materiais = "Acima da meta"
        
    if consumo_medicamentos < META_CONSUMO_MEDICAMENTOS:
        status_consumo_medicamentos = "Abaixo da meta"
    elif consumo_medicamentos == META_CONSUMO_MEDICAMENTOS:
        status_consumo_medicamentos = "Na meta"
    else:
        status_consumo_medicamentos = "Acima da meta"
    
    # Cria o registro com todos os dados e status
    record = {
        "Data": data,
        "Paciente CM": paciente_cm,
        "Status Paciente CM": status_paciente_cm,
        "Paciente CTI": paciente_cti,
        "Status Paciente CTI": status_paciente_cti,
        "Consumo Materiais": consumo_materiais,
        "Status Consumo Materiais": status_consumo_materiais,
        "Consumo Medicamentos": consumo_medicamentos,
        "Status Consumo Medicamentos": status_consumo_medicamentos
    }
    records.append(record)
    
    # Atualiza os arquivos de dados e o relatório
    save_records()
    update_report_file()
    
    messagebox.showinfo("Sucesso", f"Registro do dia {data} salvo com sucesso!")
    
    # Limpa os campos para o próximo registro
    entry_data.delete(0, tk.END)
    entry_paciente_cm.delete(0, tk.END)
    entry_paciente_cti.delete(0, tk.END)
    entry_consumo_materiais.delete(0, tk.END)
    entry_consumo_medicamentos.delete(0, tk.END)

def gerar_relatorio():
    """Abre uma janela com o relatório acumulado (usando os dados em memória)."""
    report_window = tk.Toplevel(root)
    report_window.title("Relatório Cumulativo de Indicadores")
    text = tk.Text(report_window, width=100, height=30)
    text.pack(padx=10, pady=10)
    
    if not records:
        text.insert(tk.END, "Nenhum registro disponível.")
    else:
        for rec in records:
            text.insert(tk.END, f"Data: {rec['Data']}\n")
            text.insert(tk.END, f"  Paciente dia - CM: {rec['Paciente CM']} -> {rec['Status Paciente CM']}\n")
            text.insert(tk.END, f"  Paciente dia - CTI: {rec['Paciente CTI']} -> {rec['Status Paciente CTI']}\n")
            text.insert(tk.END, f"  Consumo Estoque - Material Hospitalar: {format_currency(rec['Consumo Materiais'])} -> {rec['Status Consumo Materiais']}\n")
            text.insert(tk.END, f"  Consumo Estoque - Medicamentos: {format_currency(rec['Consumo Medicamentos'])} -> {rec['Status Consumo Medicamentos']}\n")
            text.insert(tk.END, "-" * 50 + "\n")

def gerar_relatorio_periodo():
    """Abre uma janela para filtrar o relatório por período (data inicial e final)."""
    filtro_window = tk.Toplevel(root)
    filtro_window.title("Filtrar Relatório por Período")
    
    label_dt_inicial = tk.Label(filtro_window, text="Data Inicial (dd/mm/aaaa):")
    label_dt_inicial.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    entry_dt_inicial = tk.Entry(filtro_window)
    entry_dt_inicial.grid(row=0, column=1, padx=5, pady=5)
    
    label_dt_final = tk.Label(filtro_window, text="Data Final (dd/mm/aaaa):")
    label_dt_final.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    entry_dt_final = tk.Entry(filtro_window)
    entry_dt_final.grid(row=1, column=1, padx=5, pady=5)
    
    def filtrar_relatorio():
        dt_inicial_str = entry_dt_inicial.get()
        dt_final_str = entry_dt_final.get()
        try:
            dt_inicial = datetime.strptime(dt_inicial_str, "%d/%m/%Y")
            dt_final = datetime.strptime(dt_final_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "Verifique o formato das datas. Use dd/mm/aaaa.")
            return
        
        filtered_records = []
        for rec in records:
            try:
                rec_dt = datetime.strptime(rec["Data"], "%d/%m/%Y")
            except ValueError:
                continue
            if dt_inicial <= rec_dt <= dt_final:
                filtered_records.append(rec)
        
        report_window_periodo = tk.Toplevel(filtro_window)
        report_window_periodo.title("Relatório por Período")
        text = tk.Text(report_window_periodo, width=100, height=30)
        text.pack(padx=10, pady=10)
        
        if not filtered_records:
            text.insert(tk.END, "Nenhum registro disponível no período informado.")
        else:
            for rec in filtered_records:
                text.insert(tk.END, f"Data: {rec['Data']}\n")
                text.insert(tk.END, f"  Paciente dia - CM: {rec['Paciente CM']} -> {rec['Status Paciente CM']}\n")
                text.insert(tk.END, f"  Paciente dia - CTI: {rec['Paciente CTI']} -> {rec['Status Paciente CTI']}\n")
                text.insert(tk.END, f"  Consumo Estoque - Material Hospitalar: {format_currency(rec['Consumo Materiais'])} -> {rec['Status Consumo Materiais']}\n")
                text.insert(tk.END, f"  Consumo Estoque - Medicamentos: {format_currency(rec['Consumo Medicamentos'])} -> {rec['Status Consumo Medicamentos']}\n")
                text.insert(tk.END, "-" * 50 + "\n")
    
    button_filtrar = tk.Button(filtro_window, text="Gerar Relatório", command=filtrar_relatorio)
    button_filtrar.grid(row=2, column=0, columnspan=2, pady=10)

# Janela principal
root = tk.Tk()
root.title("Sistema de Análise de Indicadores - HOSPITAL CASA SÃO BERNARDO")

# Carrega os registros salvos (se existirem)
load_records()

# Layout dos campos para entrada dos dados
label_data = tk.Label(root, text="Data (dd/mm/aaaa):")
label_data.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
entry_data = tk.Entry(root)
entry_data.grid(row=0, column=1, padx=5, pady=5)

label_paciente_cm = tk.Label(root, text="Paciente dia - CM:")
label_paciente_cm.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
entry_paciente_cm = tk.Entry(root)
entry_paciente_cm.grid(row=1, column=1, padx=5, pady=5)

label_paciente_cti = tk.Label(root, text="Paciente dia - CTI:")
label_paciente_cti.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
entry_paciente_cti = tk.Entry(root)
entry_paciente_cti.grid(row=2, column=1, padx=5, pady=5)

label_consumo_materiais = tk.Label(root, text="Consumo Estoque - Material Hospitalar (R$):")
label_consumo_materiais.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
entry_consumo_materiais = tk.Entry(root)
entry_consumo_materiais.grid(row=3, column=1, padx=5, pady=5)

label_consumo_medicamentos = tk.Label(root, text="Consumo Estoque - Medicamentos (R$):")
label_consumo_medicamentos.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
entry_consumo_medicamentos = tk.Entry(root)
entry_consumo_medicamentos.grid(row=4, column=1, padx=5, pady=5)

# Botões para salvar os registros e gerar os relatórios
button_salvar = tk.Button(root, text="Salvar Registro", command=submit_data)
button_salvar.grid(row=5, column=0, padx=5, pady=10)

button_relatorio = tk.Button(root, text="Relatório Acumulado", command=gerar_relatorio)
button_relatorio.grid(row=5, column=1, padx=5, pady=10)

button_relatorio_periodo = tk.Button(root, text="Relatório por Período", command=gerar_relatorio_periodo)
button_relatorio_periodo.grid(row=6, column=0, columnspan=2, padx=5, pady=10)

root.mainloop()
