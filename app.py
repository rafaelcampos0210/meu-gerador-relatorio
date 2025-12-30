import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

def formatar_fonte(run, tamanho=11, negrito=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE Oficial", layout="centered")

st.title("🚓 Gerador de Relatório Oficial")

with st.form("dados"):
    opj = st.text_input("OPJ:", "INTERCEPTUM")
    proc = st.text_input("Processo nº:", "0002343-02.2025.8.17.3410")
    data_hora = st.text_input("Data/Hora:", "22 de dezembro de 2025 às 14h23")
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")
    alvo = st.text_input("Alvo:", "ALEX DO CARMO CORREIA | CPF: 167.476.854-07")
    relato = st.text_area("Descrição da Ocorrência:", height=300)
    fotos = st.file_uploader("Subir Fotos", accept_multiple_files=True)
    gerar = st.form_submit_button("GERAR RELATÓRIO COM LOGO")

if gerar:
    doc = Document()
    
    # Margens do modelo oficial
    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Inches(0.5), Inches(0.5)
    sec.left_margin, sec.right_margin = Inches(0.7), Inches(0.7)

    # CABEÇALHO COM TABELA (Para o Logo e o Texto ficarem lado a lado)
    tab = doc.add_table(rows=1, cols=2)
    tab.columns[0].width = Inches(1.2)
    
    # Aqui o código procura a imagem que você subiu no GitHub
    try:
        par_logo = tab.cell(0, 0).paragraphs[0]
        run_logo = par_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.9))
    except Exception:
        tab.cell(0, 0).text = " " # Fica em branco se a imagem não estiver no GitHub

    # Texto Institucional
    p_head = tab.cell(0, 1).paragraphs[0]
    r_head = p_head.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1 - 16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    formatar_fonte(r_head, tamanho=10, negrito=True)

    # Título Central
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tit = p_tit.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    formatar_fonte(r_tit, tamanho=11, negrito=True)

    # Dados do Processo
    infos = [f"OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ): \"{opj}\"", f"PROCESSO nº {proc}", f"DATA/HORA: {data_hora}", f"LOCAL: {local}"]
    for info in infos:
        r = doc.add_paragraph().add_run(info)
        formatar_fonte(r)

    # Seções
    for titulo in ["DO ALVO E TESTEMUNHAS", "DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO"]:
        r_s = doc.add_paragraph().add_run(f"\n{titulo}")
        formatar_fonte(r_s, negrito=True)
        texto = alvo if "ALVO" in titulo else relato
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_t = p_t.add_run(texto)
        formatar_fonte(r_t)

    # Anexo de Fotos
    if fotos:
        for f in fotos:
            doc.add_page_break()
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(f, width=Inches(5.5))

    # Rodapé (Surubim)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_f = footer.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 36241974")
    formatar_fonte(r_f, tamanho=8)

    buf = io.BytesIO()
    doc.save(buf)
    st.download_button("⬇️ Baixar Relatório Fiel", buf.getvalue(), "Relatorio_Final_PCPE.docx")
