import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xml.etree.ElementTree as ET
from fpdf import FPDF
from datetime import datetime
import tempfile
import os

# --- SEGURANÇA ---
def check_password():
    if "password_correct" not in st.session_state:
        st.text_input("Senha de Acesso Executivo", type="password", 
                     on_change=lambda: st.session_state.update({"password_correct": st.session_state.password == "MV2026"}), 
                     key="password")
        return False
    return st.session_state["password_correct"]

if not check_password():
    st.stop()

# --- CONFIGURAÇÕES GRÁFICAS ---
plt.rcParams['figure.dpi'] = 300
sns.set_theme(style="whitegrid")

# --- MOTOR DE PROCESSAMENTO XML ---
def parse_project_xml(file_content):
    try:
        tree = ET.parse(file_content)
        root = tree.getroot()
        ns = '{http://schemas.microsoft.com/project}'
        
        # 1. Identificação do Gerente de Projetos (AssnOwner/StatusManager)
        # Busca o último proprietário de atribuição válido
        owners = [o.text for o in root.findall(f'.//{ns}AssnOwner') if o.text]
        gerente = owners[-1] if owners else "Não Informado"
        
        # 2. Governança e Replanejamento
        baselines = root.findall(f'.//{ns}Baseline')
        num_baselines = len(baselines)
        txt_replan = f"Sim ({num_baselines})" if num_baselines > 1 else "Não"
        
        # 3. Indicadores Financeiros
        proj_pv = float(root.find(f'.//{ns}BCWS').text or 0) if root.find(f'.//{ns}BCWS') is not None else 0
        proj_ev = float(root.find(f'.//{ns}BCWP').text or 0) if root.find(f'.//{ns}BCWP') is not None else 0
        proj_ac = float(root.find(f'.//{ns}ACWP').text or 0) if root.find(f'.//{ns}ACWP') is not None else 0
        progresso_pct = float(root.find(f'.//{ns}PercentComplete').text or 0) if root.find(f'.//{ns}PercentComplete') is not None else 0

        # 4. Auditoria PERT e Integridade
        tasks = []
        pert_errors = 0
        for task in root.findall(f'.//{ns}Task'):
            name = task.find(f'{ns}Name')
            is_summary = task.find(f'{ns}Summary')
            if name is not None and (is_summary is None or is_summary.text == '0'):
                def get_days(tag):
                    node = task.find(f'{ns}{tag}')
                    if node is not None:
                        val = node.text.replace('PT', '').replace('H', '*60+').replace('M', '*1+').replace('S', '*0').strip('+')
                        try: return max(eval(val) / 480, 0.001)
                        except: return 0.001
                    return 0.001
                ot, mp, ps = get_days('Duration1'), get_days('Duration'), get_days('Duration2')
                if ot == ps: pert_errors += 1
                tasks.append({"Otimista": ot, "Mais_Provavel": mp, "Pessimista": ps})
        
        # Cálculo do Score Dinâmico (0-10)
        score = 0
        if num_baselines > 0: score += 4
        if proj_ac > 0: score += 3
        if len(tasks) > 0 and (pert_errors / len(tasks)) < 0.2: score += 3
        
        return {
            "Gerente": gerente,
            "Replan": txt_replan,
            "PV": proj_pv, "EV": proj_ev, "AC": proj_ac,
            "Score": score, "Progresso": progresso_pct
        }
    except: return None

class ExecutivePDF(FPDF):
    def add_watermark(self):
        self.set_font("Helvetica", "B", 50)
        self.set_text_color(240, 240, 240)
        self.rotate(45, 100, 150)
        self.text(35, 190, "CONFIDENCIAL")
        self.rotate(0)
        self.set_text_color(0)

    def header_report(self, title):
        self.add_page()
        self.add_watermark()
        self.rect(5, 5, 200, 287)
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(0, 51, 102)
        self.cell(190, 15, title, ln=True, align='C')
        self.ln(5)

# --- INTERFACE ---
st.set_page_config(page_title="Portfolio Intelligence", layout="wide")
st.title("🛰️ Portfolio Intelligence: Central de Comando MV")

with st.sidebar:
    st.header("📂 Governança de Dados")
    uploaded_files = st.file_uploader("Upload Projetos (XML)", type="xml", accept_multiple_files=True)

if uploaded_files:
    results = []
    for file in uploaded_files:
        data = parse_project_xml(file)
        if data:
            spi = data["EV"] / data["PV"] if data["PV"] > 0 else (1.0 if data["Progresso"] == 100 else 0.0)
            cpi = data["EV"] / data["AC"] if data["AC"] > 0 else 1.0
            
            # Status Inteligente
            if spi >= 1.0 and cpi >= 1.0 and data["Progresso"] == 100:
                status = "Concluído com sucesso"
            else:
                status = "CRÍTICO" if spi < 0.8 else ("ALERTA" if spi < 0.9 else "SAUDÁVEL")

            results.append({
                "Projeto": file.name.replace('.xml', '').upper(),
                "Gerente de Projetos": data["Gerente"],
                "SPI": round(spi, 2),
                "CPI": round(cpi, 2),
                "Investimento Recuperável": max(0, data["PV"] - data["EV"]),
                "Score Qualidade": f"{data['Score']}/10",
                "Replan?": data["Replan"],
                "Status": status
            })

    df_port = pd.DataFrame(results)

    # --- PAINEL CONSOLIDADO ---
    st.subheader("📋 Painel de Controle Consolidado")
    st.dataframe(df_port[['Projeto', 'Gerente de Projetos', 'SPI', 'CPI', 'Investimento Recuperável', 'Score Qualidade', 'Status']], 
                 use_container_width=True)

    if st.button("🚀 Gerar RELATORIO_MV_DIRETORIA"):
        pdf = ExecutivePDF()
        pdf.header_report("RELATÓRIO DE PERFORMANCE E AUDITORIA")
        
        # 1. Visão Geral
        pdf.set_fill_color(230, 230, 230); pdf.set_font("Helvetica", "B", 11)
        pdf.cell(190, 10, " 1. SUMÁRIO EXECUTIVO DO PORTFÓLIO", ln=True, fill=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(95, 10, f" Total de Projetos: {len(df_port)}", border='L')
        pdf.cell(95, 10, f" Investimento Recuperável Total: R$ {df_port['Investimento Recuperável'].sum():,.2f}", ln=True, border='R')

        # 2. Tabela Master
        pdf.ln(5); pdf.set_font("Helvetica", "B", 11)
        pdf.cell(190, 10, " 2. DETALHAMENTO POR UNIDADE E GERENTE", ln=True, fill=True)
        
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(50, 8, " Projeto", border=1, fill=True)
        pdf.cell(40, 8, " Gerente", border=1, fill=True)
        pdf.cell(15, 8, " SPI", border=1, align='C', fill=True)
        pdf.cell(15, 8, " CPI", border=1, align='C', fill=True)
        pdf.cell(30, 8, " Invest. Recup.", border=1, align='C', fill=True)
        pdf.cell(15, 8, " Score", border=1, align='C', fill=True)
        pdf.cell(25, 8, " Status", border=1, ln=True, align='C', fill=True)
        
        pdf.set_font("Helvetica", "", 6)
        for _, row in df_port.iterrows():
            pdf.cell(50, 7, f" {row['Projeto'][:28]}", border=1)
            pdf.cell(40, 7, f" {row['Gerente de Projetos'][:22]}", border=1)
            pdf.cell(15, 7, f" {row['SPI']:.2f}", border=1, align='C')
            pdf.cell(15, 7, f" {row['CPI']:.2f}", border=1, align='C')
            pdf.cell(30, 7, f" R$ {row['Investimento Recuperável']:,.2f}", border=1, align='R')
            pdf.cell(15, 7, f" {row['Score Qualidade']}", border=1, align='C')
            
            # Cores de Status no PDF
            if row['Status'] == "Concluído com sucesso": pdf.set_fill_color(200, 255, 200)
            elif row['Status'] == "CRÍTICO": pdf.set_fill_color(255, 200, 200)
            else: pdf.set_fill_color(255, 255, 255)
            
            pdf.cell(25, 7, f" {row['Status']}", border=1, ln=True, align='C', fill=True)
            pdf.set_fill_color(255, 255, 255)

        # 3. Conclusão da Entrega
        pdf.ln(10); pdf.set_font("Helvetica", "B", 11); pdf.cell(190, 10, " 3. CONCLUSÃO E PARECER TÉCNICO", ln=True, fill=True)
        pdf.set_font("Helvetica", "", 10)
        
        txt_conclusao = (f"A auditoria consolidada identificou um volume de investimento recuperável de "
                         f"R$ {df_port['Investimento Recuperável'].sum():,.2f}. Projetos com status 'Concluído com sucesso' "
                         f"demonstram aderência total à metodologia MV. Unidades em estado 'CRÍTICO' sob gestão dos "
                         f"gerentes listados requerem plano de recuperação imediato.")
        pdf.multi_cell(190, 7, txt_conclusao)

        # Assinaturas
        pdf.set_y(250); pdf.line(20, 260, 90, 260); pdf.line(120, 260, 190, 260)
        pdf.set_font("Helvetica", "B", 10); pdf.set_y(261)
        pdf.set_x(20); pdf.cell(70, 7, "Diretoria de Operações", align='C')
        pdf.set_x(120); pdf.cell(70, 7, "Diretor de PMO / Auditor", align='C')

        st.download_button("📥 Baixar RELATORIO_MV_DIRETORIA.pdf", bytes(pdf.output()), "RELATORIO_MV_DIRETORIA.pdf")
