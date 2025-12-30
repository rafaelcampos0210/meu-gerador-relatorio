import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
from datetime import datetime

def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.bold = bold

st.set_page_config(page_title="Gerador PCPE", layout="centered")

st.title("🚓 Gerador de Relatório de Busca e Apreensão")
st.info("Este modelo segue a formatação oficial da Polícia Civil de Pernambuco.")

# --- ENTRADA DE DADOS ---
with st.expander("Informações do Processo", expanded=True):
    processo = st.text_input("Nº do Processo:", value="0002343-02.2025.8.17.3410")
    opj = st.text_input("Operação (OPJ):", value="INTERCEPTUM")
    local = st.text_input("Local da Diligência:", value="Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

with st.expander("Dados do Alvo", expanded=True):
    alvo_nome = st.text_input("Nome do Alvo:", value="ALEX DO CARMO CORREIA")
    alvo_docs = st.text_input("CPF/RG:", value="CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
    testemunha = st.text_input("Testemunha:", value="Sra. Marilene Lima do Carmo Correia (Genitora)")

with st.expander("Conteúdo do Relatório", expanded=True):
    diligencia_texto = st.text_area("Descrição da Diligência:", height=150)
    objetos_texto = st.text_area("Objetos Localizados:", height=100)
    fotos = st.file_uploader("Imagens da Diligência", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])

if st.button("🚀 GERAR DOCUMENTO FORMATADO"):
    doc = Document()

    # --- CONFIGURAÇÃO DE MARGENS ---
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # --- CABEÇALHO ---
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1 - 16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    set_font(run, 'Arial', 10, bold=True)

    # --- TÍTULO DO RELATÓRIO ---
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    set_font(run, 'Arial', 11, bold=True)

    # --- INFO BOX (OPJ e PROCESSO) ---
    doc.add_paragraph(f"OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ): \"{opj}\"")
    doc.add_paragraph(f"PROCESSO nº {processo}")
    doc.add_paragraph(f"DATA: {datetime.now().strftime('%d de dezembro de 2025')}")
    doc.add_paragraph(f"LOCAL: {local}")

    # --- SEÇÃO: ALVO ---
    h1 = doc.add_heading('DO ALVO E TESTEMUNHAS', level=1)
    doc.add_paragraph(f"ALVO: {alvo_nome} | {alvo_docs}")
    doc.add_paragraph(f"TESTEMUNHA: {testemunha}")

    # --- SEÇÃO: DILIGÊNCIA ---
    doc.add_heading('DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO', level=1)
    p_dil = doc.add_paragraph(diligencia_texto)
    p_dil.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # --- OBJETOS ---
    if objetos_texto:
        doc.add_heading('OBJETOS LOCALIZADOS', level=2)
        doc.add_paragraph(objetos_texto)

    # --- FOTOS (FORMATADAS) ---
    if fotos:
        for foto in fotos:
            doc.add_paragraph("\n") # Espaçamento
            doc.add_picture(foto, width=Inches(5))
            last_p = doc.paragraphs[-1]
            last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption = doc.add_paragraph(f"Registro Fotográfico: {foto.name}")
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- RODAPÉ E ASSINATURA ---
    doc.add_paragraph("\n\n")
    sig = doc.add_paragraph("__________________________________________")
    sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_sig = doc.add_paragraph("Assinatura do Responsável")
    name_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Gerar arquivo
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)

    st.success("✅ Relatório formatado com sucesso!")
    st.download_button("⬇️ Baixar Relatório PCPE", target, "Relatorio_Busca.docx")
