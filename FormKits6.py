import streamlit as st
import pandas as pd
import os
import requests
from fpdf import FPDF
import tempfile

# --- Configurações ---
FOLDER_PATH = r"L:\\Planejamento\\Planejamento Geral\\18 - PMO\\Projeto Kits"
ARQ_CAD_KITS = os.path.join(FOLDER_PATH, "Cadastro Kits.xlsx")
ARQ_MAT_KITS = os.path.join(FOLDER_PATH, "Materias Kits.xlsx")
LOGO_URL = "https://www.contrex.com.br/wp-content/themes/contrex_2019/images/logo_footer.png"
LOGO_PATH = os.path.join(FOLDER_PATH, "logo_footer.png")

# --- Funções ---

def baixar_logo():
    if not os.path.exists(LOGO_PATH):
        response = requests.get(LOGO_URL)
        if response.status_code == 200:
            with open(LOGO_PATH, "wb") as f:
                f.write(response.content)

def format_num(valor):
    if pd.isna(valor):
        return "0"
    s = f"{valor:.1f}"
    if s.endswith(".0"):
        s = s[:-2]
    return s.replace('.', ',')

def normalize_str(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()

def formatar_descricao_kit(row):
    tipo = row.get("Tipo de Kit", "")
    if pd.isna(tipo):
        return None
    altura = format_num(row.get('Altura'))
    largura = format_num(row.get('Largura'))
    comprimento = format_num(row.get('Comprimento'))
    descricao = f"{tipo} {altura}m x {largura}m x {comprimento}m"
    return normalize_str(descricao)

def obter_materiais(codigo, mat_kits):
    if pd.isna(codigo):
        return None
    materiais = mat_kits[mat_kits['Código'] == codigo]
    if materiais.empty:
        return None
    return materiais

def gerar_pdf(sgs, cad_kits, mat_kits, caminho_pdf):
    # Preparar colunas
    sgs["Descrição do Kit"] = sgs.apply(formatar_descricao_kit, axis=1)
    cad_kits["Descrição Kit"] = cad_kits["Descrição Kit"].apply(normalize_str)

    # Merge para trazer o código do kit
    merged = pd.merge(
        sgs,
        cad_kits[['Código', 'Descrição Kit']],
        left_on='Descrição do Kit',
        right_on='Descrição Kit',
        how='left'
    )

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)

    placas = merged['Placa'].dropna().unique()

    for placa in placas:
        dados_placa = merged[merged['Placa'] == placa]
        if dados_placa.empty:
            continue

        pdf.add_page()
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=160, y=8, w=20)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 10, f"Formulário de Materiais - Placa: {placa}", ln=True)

        for _, row in dados_placa.iterrows():
            descricao_kit = row['Descrição do Kit']
            codigo_kit = row['Código']

            # Cabeçalho com dados da OS
            pdf.set_font("Arial", size=9)
            pdf.cell(0, 6, f"Kit: {descricao_kit}", ln=True)
            pdf.cell(0, 5, f"Número OS: {row.get('Número OS', '')}", ln=True)
            pdf.cell(0, 5, f"Área: {row.get('Área', '')}", ln=True)
            pdf.cell(0, 5, f"Descrição: {row.get('Descrição', '')}", ln=True)
            pdf.cell(0, 5, f"Levantado por: {row.get('Executante', '')}  Em: {row.get('Data Execução', '')}", ln=True)
            pdf.cell(0, 5, f"Código do Kit: {codigo_kit}", ln=True)

            materiais = obter_materiais(codigo_kit, mat_kits)
            if materiais is None:
                pdf.cell(0, 6, "Kit não cadastrado ou sem materiais.", ln=True)
            else:
                fonte_tabela = 8 if len(materiais) > 10 else 9

                largura_id = 25
                largura_desc = 95
                largura_qtd = 40

                pdf.set_font("Arial", 'B', fonte_tabela)
                pdf.cell(largura_id, 5, "ID", border=1)
                pdf.cell(largura_desc, 5, "Descrição", border=1)
                pdf.cell(largura_qtd, 5, "Quantidade", border=1, ln=True)
                pdf.set_font("Arial", size=fonte_tabela)

                for _, mat in materiais.iterrows():
                    y_before = pdf.get_y()
                    x_before = pdf.get_x()

                    pdf.set_xy(x_before + largura_id, y_before)
                    y_before_desc = pdf.get_y()
                    pdf.multi_cell(largura_desc, 5, str(mat.get('Descrição', '')), border=1)
                    y_after_desc = pdf.get_y()
                    altura = y_after_desc - y_before_desc

                    pdf.set_xy(x_before, y_before)
                    pdf.cell(largura_id, altura, str(mat.get('ID', '')), border=1)

                    pdf.set_xy(x_before + largura_id + largura_desc, y_before)
                    pdf.cell(largura_qtd, altura, str(mat.get('Quantidade', '')), border=1, ln=True)

            pdf.ln(3)

    pdf.output(caminho_pdf)


# --- Streamlit Interface ---

st.title("Formulário de Materiais - Projeto Kits")

# Upload do arquivo SGS.xlsx pelo usuário
uploaded_file = st.file_uploader("Selecione o arquivo SGS.xlsx", type=["xlsx"])

if uploaded_file is not None:
    try:
        sgs = pd.read_excel(uploaded_file)
        cad_kits = pd.read_excel(ARQ_CAD_KITS)
        mat_kits = pd.read_excel(ARQ_MAT_KITS)

        # Verificar se colunas obrigatórias existem no SGS
        colunas_esperadas = ["Número OS", "Área", "Descrição", "Executante", "Placa",
                             "Tipo de Kit", "Altura", "Largura", "Comprimento", "Data Execução"]
        faltantes = [c for c in colunas_esperadas if c not in sgs.columns]
        if faltantes:
            st.error(f"Arquivo SGS.xlsx não contém as colunas necessárias: {faltantes}")
        else:
            baixar_logo()

            if st.button("📄 Gerar PDF"):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                        gerar_pdf(sgs, cad_kits, mat_kits, tmp_pdf.name)
                        tmp_pdf_path = tmp_pdf.name

                    with open(tmp_pdf_path, "rb") as f:
                        pdf_bytes = f.read()

                    st.success("✅ PDF gerado com sucesso! Faça o download abaixo:")
                    st.download_button(
                        label="⬇️ Baixar PDF",
                        data=pdf_bytes,
                        file_name="Relatorio_Materiais.pdf",
                        mime="application/pdf"
                    )

                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("🔁 Novo Formulário"):
                            st.experimental_rerun()
                    with col2:
                        if st.button("❌ Finalizar"):
                            st.stop()

                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {e}")

    except Exception as e:
        st.error(f"Erro ao ler arquivos: {e}")

else:
    st.info("Por favor, envie o arquivo SGS.xlsx para começar.")

