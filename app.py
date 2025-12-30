import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# Função para garantir que a fonte seja Arial 11 (Padrão PCPE)
def aplicar_estilo_oficial(run, tamanho=11, negrito=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE Oficial", layout="centered")

st.title("🚓 Gerador de Relatório (Modelo Fiel)")

with st.form("formulario_pcpe"):
    st.subheader("Dados do Cabeçalho")
    opj = st.text_input("OPJ:", "INTERCEPTUM")
    processo = st.text_input("Nº Processo:", "0002343-02.2025.8.17.3410")
    data_extenso = st.text_input("Data:", "22 de dezembro de 2025")
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")
    
    st.subheader("Alvo e Relato")
    alvo = st.text_input("Alvo:", "ALEX DO CARMO CORREIA | CPF: 167.476.854-07")
    nascimento = st.text_input("Nascimento:", "15/04/2004")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")
    
    relato = st.text_area("Descrição da Diligência:", height=300)
    fotos = st.file_uploader("Anexar Imagens", accept_multiple_files=True)
    
    botao = st.form_submit_button("GERAR RELATÓRIO IDÊNTICO AO MODELO")

if botao:
    doc = Document()
    
    # MARGENS: Ajuste para o padrão do modelo (Estreitas)
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    # CABEÇALHO COM TABELA (Para o logo ficar à esquerda do texto)
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.2)
    table.columns[1].width = Inches(5.5)
    
    # Coluna do Logo
    try:
        run_logo = table.cell(0, 0).paragraphs[0].add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.85))
    except:
        table.cell(0, 0).text = " "

    # Coluna do Texto Institucional
    p_head = table.cell(0, 1).paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_head = p_head.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo_oficial(run_head, tamanho=10, negrito=True)

    # TÍTULO CENTRALIZADO
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = p_titulo.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    aplicar_estilo_oficial(run_titulo, tamanho=12, negrito=True)

    # BLOCO DE INFORMAÇÕES (OPJ, PROCESSO, ETC)
    def add_linha_info(label, valor):
        p = doc.add_paragraph()
        run_label = p.add_run(f"{label}: ")
        aplicar_estilo_oficial(run_label, negrito=True)
        run_valor = p.add_run(valor)
        aplicar_estilo_oficial(run_valor)

    add_linha_info("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_linha_info("PROCESSO nº", processo)
    add_linha_info("DATA", data_extenso)
    add_linha_info("LOCAL", local)

    # SEÇÃO ALVOS
    doc.add_paragraph()
    run_s1 = doc.add_paragraph().add_run("DO ALVO E TESTEMUNHAS")
    aplicar_estilo_oficial(run_s1, negrito=True)
    
    p_alvo = doc.add_paragraph()
    run_alvo = p_alvo.add_run(f"ALVO: {alvo}\nNascimento: {nascimento}\nADVOGADO: {advogado}\nTESTEMUNHA: {testemunha}")
    aplicar_estilo_oficial(run_alvo)

    # SEÇÃO DILIGÊNCIA
    run_s2 = doc.add_paragraph().add_run("\nDA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    aplicar_estilo_oficial(run_s2, negrito=True)
    
    p_relato = doc.add_paragraph()
    p_relato.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_relato = p_relato.add_run(relato)
    aplicar_estilo_oficial(run_relato)

    # FOTOS (Uma por página como no modelo)
    if fotos:
        for foto in fotos:
            doc.add_page_break()
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(foto, width=Inches(5.8))
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_cap = p_cap.add_run(f"Registro Fotográfico: {foto.name}")
            aplicar_estilo_oficial(run_cap, tamanho=10)

    # RODAPÉ
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = footer.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 36241974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    aplicar_estilo_oficial(run_foot, tamanho=8)

    # Salvar para download
    buf = io.BytesIO()
    doc.save(buf)
    st.download_button("⬇️ BAIXAR RELATÓRIO CONFIGURADO", buf.getvalue(), "Relatorio_Oficial.docx")
