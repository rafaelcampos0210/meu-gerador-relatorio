import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

def format_font(run, size=11, bold=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(size)
    run.bold = bold

st.set_page_config(page_title="Gerador PCPE", layout="centered")
st.title("🚓 Gerador de Relatório Oficial")

with st.form("form_pcpe"):
    opj = st.text_input("OPJ:", "INTERCEPTUM")
    processo = st.text_input("Processo:", "0002343-02.2025.8.17.3410")
    data_hora = st.text_input("Data/Hora:", "22 de dezembro de 2025 às 14h23")
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")
    alvo = st.text_input("Alvo:", "ALEX DO CARMO CORREIA | CPF: 167.476.854-07")
    relato = st.text_area("Descrição da Diligência:", height=200)
    fotos = st.file_uploader("Fotos", accept_multiple_files=True)
    enviar = st.form_submit_button("Gerar Relatório Idêntico")

if enviar:
    doc = Document()
    # Ajuste de margens conforme o DOCX original
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Inches(0.5), Inches(0.5)
    section.left_margin, section.right_margin = Inches(0.7), Inches(0.7)

    # CABEÇALHO (Tabela para alinhar Logo e Texto)
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.0)
    try:
        run_img = table.cell(0,0).paragraphs[0].add_run()
        run_img.add_picture('logo_pc.png', width=Inches(0.85))
    except: pass
    
    txt_head = table.cell(0,1).paragraphs[0]
    run_h = txt_head.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    format_font(run_h, size=10, bold=True)

    # TÍTULO CENTRALIZADO
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_t.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    format_font(run_t, size=11, bold=True)

    # CORPO DO TEXTO (Igual ao PDF/DOCX)
    infos = [f"OPJ: \"{opj}\"", f"PROCESSO nº {processo}", f"DATA/HORA: {data_hora}", f"LOCAL: {local}"]
    for info in infos:
        run_i = doc.add_paragraph().add_run(info)
        format_font(run_i)

    # SEÇÕES
    for sec in ["DO ALVO E TESTEMUNHAS", "DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO"]:
        run_s = doc.add_paragraph().add_run(f"\n{sec}")
        format_font(run_s, bold=True)
        txt = alvo if "ALVO" in sec else relato
        run_txt = doc.add_paragraph().add_run(txt)
        format_font(run_txt)
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # FOTOS E RODAPÉ
    if fotos:
        for f in fotos:
            doc.add_page_break()
            doc.add_picture(f, width=Inches(5.5))
    
    # RODAPÉ OFICIAL
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = footer.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 36241974")
    format_font(run_f, size=8)

    output = io.BytesIO()
    doc.save(output)
    st.download_button("Baixar Relatório", output.getvalue(), "Relatorio_Oficial.docx")
