import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
from datetime import datetime

# Função para configurar fonte padrão
def configurar_fonte(run, nome='Arial', tamanho=11, negrito=False):
    run.font.name = nome
    run._element.rPr.rFonts.set(qn('w:eastAsia'), nome)
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE", layout="centered")

st.title("🚓 Gerador de Relatório Oficial - PCPE")

# --- ENTRADA DE DADOS ---
with st.expander("Dados do Cabeçalho", expanded=True):
    opj = st.text_input("OPJ:", value="INTERCEPTUM")
    processo = st.text_input("Processo nº:", value="0002343-02.2025.8.17.3410")
    data_hora = st.text_input("Data e Hora:", value="22 de dezembro de 2025 às 14h23")
    local = st.text_input("Local:", value="Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

with st.expander("Alvo e Testemunhas", expanded=True):
    alvo = st.text_input("Dados do Alvo:", value="ALEX DO CARMO CORREIA | CPF: 167.476.854-07")
    testemunha = st.text_input("Testemunha:", value="Sra. Marilene Lima do Carmo Correia (Genitora)")

relato = st.text_area("Descrição da Ocorrência (Diligência):", height=300)
fotos = st.file_uploader("Imagens da Ocorrência", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

if st.button("🚀 Gerar Relatório no Formato PCPE"):
    doc = Document()
    
    # --- CABEÇALHO OFICIAL ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    configurar_fonte(run, tamanho=10, negrito=True)

    # --- TÍTULO DO RELATÓRIO ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    configurar_fonte(run, tamanho=12, negrito=True)

    # --- INFO BOX ---
    doc.add_paragraph(f"OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ): \"{opj}\"")
    doc.add_paragraph(f"PROCESSO nº {processo}")
    doc.add_paragraph(f"DATA/HORA: {data_hora}")
    doc.add_paragraph(f"LOCAL: {local}")

    # --- SEÇÃO 1: ALVOS ---
    p = doc.add_heading('DO ALVO E TESTEMUNHAS', level=1)
    doc.add_paragraph(f"ALVO: {alvo}")
    doc.add_paragraph(f"TESTEMUNHA: {testemunha}")

    # --- SEÇÃO 2: DILIGÊNCIA ---
    doc.add_heading('DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO', level=1)
    p = doc.add_paragraph(relato)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # --- FOTOS ---
    if fotos:
        doc.add_heading('ANEXO FOTOGRÁFICO', level=1)
        for i, foto in enumerate(fotos):
            doc.add_picture(foto, width=Inches(5.5))
            p_foto = doc.paragraphs[-1]
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            legenda = doc.add_paragraph(f"Registro Fotográfico {i+1}")
            legenda.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_page_break()

    # --- ASSINATURA ---
    doc.add_paragraph("\n\n")
    p = doc.add_paragraph("__________________________________________")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph("RAFAEL DE ALBUQUERQUE CAMPOS\nInvestigador de Polícia")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Salvar
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    
    st.download_button(label="⬇️ Baixar Relatório PCPE", data=target, file_name="Relatorio_PCPE.docx")
