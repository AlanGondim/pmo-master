import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xml.etree.ElementTree as ET
import numpy_financial as npf
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

# --- PROCESSAMENTO XML COM AUDITORIA ---
def parse_project_xml(file_content):
    try:
        tree = ET.parse(file_content)
        root = tree.getroot()
        ns = '{http://schemas.microsoft.com/project}'
        
        # 1. Auditoria de Governança (Linhas de Base)
        baselines = root.findall(f'.//{ns}Baseline')
        num_baselines = len(baselines)
        baseline_saved = num_baselines > 0
        
        proj_pv = float(root.find(f'.//{ns}BCWS').text or 0) if root.find(f'.//{ns}BCWS') is not None else 0
        proj_ev = float(root.find(f'.//{ns}BCWP').text or 0) if root.find(f'.//{ns}BCWP') is not None else 0
        proj_ac = float(root.find(f'.//{ns}ACWP').text or 0) if root.find(f'.//{ns}ACWP') is not None else 0

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
                
                tasks.append({
                    "Otimista": ot, "Mais_Provavel": mp, "Pessimista": ps,
                    "Custo_Fixo": float(task.find(f'{ns}Cost').text or 0) if task.find(f'{ns}Cost') is not None else 0
                })
        
        # 2. Score de Integridade (0-10)
        score = 0
        if baseline_saved: score += 4
        if proj_ac > 0: score += 3
        if len(tasks) > 0 and (pert_errors / len(tasks)) < 0.2: score += 3
        
        df = pd.DataFrame(tasks)
        if not df.empty:
            df['Otimista_F'] = df[['Otimista', 'Mais_Provavel', 'Pessimista']].min(axis=1)
            df['Pessimista_F'] = df[['Otimista', 'Mais_Provavel', 'Pessimista']].max(axis=1)
            df['Mais_Provavel_F'] = df[['Otimista', 'Mais_Provavel', 'Pessimista']].median(axis=1)
            mask = df['Otimista_F'] == df['Pessimista_F']
            df.loc[mask, 'Pessimista_F'] = df.loc[mask, 'Pessimista_F'] + 0.01
            
        return df, proj_pv, proj_ev, proj_ac, score, num_baselines
    except: return pd.DataFrame(), 0, 0, 0, 0, 0

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
        self.set_font("Helvetica", "I", 9)
        self.cell(190, 5, f"Auditado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True, align='C')
        self.ln(5)

# --- INTERFACE ---
st.set_page_config(page_title="Portfolio Intelligence", layout="wide")
st.title("🛰️ Portfolio Intelligence: Auditoria & Benchmarking de Programas & Projetos")

with st.sidebar:
    st.header("📂 Governança de Dados")
    uploaded_files = st.file_uploader("Upload Projetos (XML)", type="xml", accept_multiple_files=True)
    estresse = st.slider("Stress Test Operacional (%)", 0, 100, 30) / 100
    burn_rate_base = st.number_input("Custo Diário Base (R$)", value=1800)
    burn_rate = burn_rate_base * (1 + estresse)

if uploaded_files:
    results = []
    for file in uploaded_files:
        df, pv, ev, ac, score, n_base = parse_project_xml(file)
        if not df.empty:
            spi = ev / pv if pv > 0 else 0.0
            cpi = ev / ac if ac > 0 else 1.0
            inv_rec = max(0, pv - ev)
            status_pdf = "CRITICO" if spi < 0.8 or score < 5 else ("ALERTA" if spi < 0.9 or score < 8 else "SAUDAVEL")
            
            results.append({
                "Projeto": file.name.replace('.xml', '').upper(),
                "SPI": spi, "CPI": cpi, "Score": score, 
                "Linhas_Base": n_base, "Invest_Recup": inv_rec, "Status": status_pdf
            })

    df_port = pd.DataFrame(results)

    # --- SEÇÃO DE BENCHMARKING ---
    st.subheader("📊 Benchmarking de Performance vs Governança")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Ranking por Eficiência (SPI)**")
        fig_spi, ax_spi = plt.subplots(figsize=(10, 6))
        sns.barplot(data=df_port.sort_values("SPI", ascending=False), x="SPI", y="Projeto", palette="viridis", ax=ax_spi)
        ax_spi.axvline(0.95, color="red", linestyle="--")
        st.pyplot(fig_spi)

    with col2:
        st.write("**Ranking por Qualidade de Dados (Audit Score)**")
        fig_score, ax_score = plt.subplots(figsize=(10, 6))
        sns.barplot(data=df_port.sort_values("Score", ascending=False), x="Score", y="Projeto", palette="magma", ax=ax_score)
        st.pyplot(fig_score)
    
    

    st.subheader("📋 Painel de Controle Consolidado")
    st.dataframe(df_port.style.highlight_max(axis=0, subset=['Score']), use_container_width=True)

    if st.button("🚀 Gerar Relatório Master para Diretoria"):
        pdf = ExecutivePDF()
        pdf.header_report("RELATÓRIO CONSOLIDADO DE GOVERNANÇA")
        
        # 1. Sumário do Portfólio
        pdf.set_fill_color(230, 230, 230); pdf.set_font("Helvetica", "B", 11)
        pdf.cell(190, 10, " 1. VISÃO GERAL DO ECOSSISTEMA", ln=True, fill=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(95, 10, f" Total de Projetos: {len(df_port)}", border='L')
        pdf.cell(95, 10, f" SPI Médio: {df_port['SPI'].mean():.2f}", ln=True, border='R')
        pdf.cell(190, 10, f" Investimento Total Recuperável: R$ {df_port['Invest_Recup'].sum():,.2f}", border='LRB', ln=True)
        
        # 2. Tabela Master
        pdf.ln(5); pdf.set_font("Helvetica", "B", 11)
        pdf.cell(190, 10, " 2. RANKING DE AUDITORIA E PERFORMANCE", ln=True, fill=True)
        
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(65, 8, " Projeto", border=1, fill=True)
        pdf.cell(20, 8, " SPI", border=1, align='C', fill=True)
        pdf.cell(20, 8, " CPI", border=1, align='C', fill=True)
        pdf.cell(25, 8, " Score (0-10)", border=1, align='C', fill=True)
        pdf.cell(25, 8, " L. Base", border=1, align='C', fill=True)
        pdf.cell(35, 8, " Status", border=1, ln=True, align='C', fill=True)
        
        pdf.set_font("Helvetica", "", 7)
        for _, row in df_port.iterrows():
            pdf.cell(65, 7, f" {row['Projeto'][:35]}", border=1)
            pdf.cell(20, 7, f" {row['SPI']:.2f}", border=1, align='C')
            pdf.cell(20, 7, f" {row['CPI']:.2f}", border=1, align='C')
            pdf.cell(25, 7, f" {row['Score']}/10", border=1, align='C')
            pdf.cell(25, 7, f" {row['Linhas_Base']}", border=1, align='C')
            
            if row['Status'] == "CRITICO": pdf.set_fill_color(255, 200, 200)
            elif row['Status'] == "ALERTA": pdf.set_fill_color(255, 255, 200)
            else: pdf.set_fill_color(200, 255, 200)
            
            pdf.cell(35, 7, f" {row['Status']}", border=1, ln=True, align='C', fill=True)
            pdf.set_fill_color(255, 255, 255)

        # 3. Gráficos de Benchmarking no PDF
        pdf.ln(5)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_spi:
            fig_spi.savefig(tmp_spi.name, bbox_inches='tight')
            pdf.image(tmp_spi.name, x=15, y=pdf.get_y(), w=85)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_score:
            fig_score.savefig(tmp_score.name, bbox_inches='tight')
            pdf.image(tmp_score.name, x=105, y=pdf.get_y(), w=85)

        # 4. Parecer Técnico
        pdf.set_y(pdf.get_y() + 60)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(190, 10, " 3. CONCLUSÃO E PARECER DA DIRETORIA", ln=True, fill=True)
        pdf.set_font("Helvetica", "", 10)
        txt = (f"Foram identificados {len(df_port[df_port['Score'] < 6])} projetos com baixa integridade de dados. "
               f"O volume de investimento recuperável consolidado (R$ {df_port['Invest_Recup'].sum():,.2f}) "
               "exige revisão imediata das unidades em estado CRÍTICO.")
        pdf.multi_cell(190, 7, txt)

        # Assinaturas
        pdf.set_y(245); pdf.line(20, 260, 90, 260); pdf.line(120, 260, 190, 260)
        pdf.set_font("Helvetica", "B", 10); pdf.set_y(261)
        pdf.set_x(20); pdf.cell(70, 7, "Diretor de PMO / Auditor", align='C')
        pdf.set_x(120); pdf.cell(70, 7, "Diretor de Operações", align='C')

        st.download_button("📥 Baixar Relatório Master Consolidado", bytes(pdf.output()), "AUDITORIA_MASTER.pdf")
