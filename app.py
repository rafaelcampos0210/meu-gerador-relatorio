import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
from datetime import datetime

# Configuração de fonte Arial (Padrão PCPE)
def aplicar_estilo_pcpe(run, tamanho=11, negrito=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE Oficial", layout="centered")

st.title("🚓 Gerador de Relatório - Padrão Fiel ao DOCX")

# --- ENTRADA DE DADOS ---
with st.form("dados_relatorio"):
    col1, col2 = st.columns(2)
    with col1:
        opj = st.text_input("OPJ:", "INTERCEPTUM")
        processo = st.text_input("Processo nº:", "0002343-02.2025.8.17.3410")
    with col2:
        data_extenso = st.text_input("Data:", "22 de dezembro de 2025")
        hora = st.text_input("Hora:", "14h23")
    
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")
    
    st.markdown("---")
    alvo_nome = st.text_input("Nome do Alvo:", "ALEX DO CARMO CORREIA")
    alvo_detalhes = st.text_input("Documentos/Nascimento:", "CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")
    
    st.markdown("---")
    relato = st.text_area("Texto da Diligência:", height=300)
    conclusao = st.text_area("Conclusão:", "A diligência transcorreu sem intercorrências...")
    
    fotos = st.file_uploader("Imagens da Ocorrência", accept_multiple_files=True)
    
    gerar = st.form_submit_button("🚀 GERAR RELATÓRIO IDÊNTICO")

if gerar:
    doc = Document()
    
    # Configuração de Margens Estreitas (conforme o DOCX enviado)
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # --- CABEÇALHO COM LOGO (TABELA INVISÍVEL) ---
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.2)
    table.columns[1].width = Inches(5.0)
    
    # Coluna 1: Logo
    try:
        run_logo = table.cell(0, 0).paragraphs[0].add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.9))
    except:
        table.cell(0, 0).text = " "

    # Coluna 2: Texto Institucional
    p_head = table.cell(0, 1).paragraphs[0]
    run_head = p_head.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo_pcpe(run_head, tamanho=10, negrito=True)

    # --- TÍTULO ---
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titulo = p_titulo.add_run("\nRELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    aplicar_estilo_pcpe(run_titulo, tamanho=11, negrito=True)

    # --- INFOS INICIAIS ---
    def add_info(label, text):
        p = doc.add_paragraph()
        run_l = p.add_run(f"{label}: ")
        aplicar_estilo_pcpe(run_l, negrito=True)
        run_t = p.add_run(text)
        aplicar_estilo_pcpe(run_t)

    add_info("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_info("PROCESSO nº", processo)
    add_info("DATA", data_extenso)
    add_info("HORA", hora)
    add_info("LOCAL", local)

    # --- SEÇÃO ALVO ---
    p_alvo_h = doc.add_paragraph()
    run_alvo_h = p_alvo_h.add_run("\nDO ALVO E TESTEMUNHAS")
    aplicar_estilo_pcpe(run_alvo_h, negrito=True)

    p_alvo_d = doc.add_paragraph()
    run_alvo_d = p_alvo_d.add_run(f"ALVO: {alvo_nome} | {alvo_detalhes}\nADVOGADO: {advogado}\nTESTEMUNHA: {testemunha}")
    aplicar_estilo_pcpe(run_alvo_d)

    # --- SEÇÃO DILIGÊNCIA ---
    p_dil_h = doc.add_paragraph()
    run_dil_h = p_dil_h.add_run("\nDA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    aplicar_estilo_pcpe(run_dil_h, negrito=True)

    p_relato = doc.add_paragraph()
    p_relato.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_relato = p_relato.add_run(relato)
    aplicar_estilo_pcpe(run_relato)

    # --- FOTOS (FORMATO ANEXO) ---
    if fotos:
        for foto in fotos:
            doc.add_page_break()
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(foto, width=Inches(5.5))
            cap = doc.add_paragraph(f"Registro Fotográfico - {foto.name}")
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- CONCLUSÃO ---
    p_conc_h = doc.add_paragraph()
    run_conc_h = p_conc_h.add_run("\nCONCLUSÃO")
    aplicar_estilo_pcpe(run_conc_h, negrito=True)
    
    p_conc_t = doc.add_paragraph()
    p_conc_t.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run_conc_t = p_conc_t.add_run(conclusao)
    aplicar_estilo_pcpe(run_conc_t)

    # --- ASSINATURA ---
    doc.add_paragraph("\n\n")
    p_ass = doc.add_paragraph()
    p_ass.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_ass = p_ass.add_run("__________________________________________\nInvestigador de Polícia")
    aplicar_estilo_pcpe(run_ass)

    # --- RODAPÉ ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = p_foot.add_run("Av. São Sebastião - São Sebastião, Surubim - PE, 55750-000\nFone: (81) 36241974 | WhatsApp +55 81 99488-7096\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    aplicar_estilo_pcpe(run_foot, tamanho=8)

    # Gerar Download
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    
    st.success("✅ Relatório formatado conforme o modelo!")
    st.download_button("⬇️ Baixar Relatório Fiel", target, f"Relatorio_{alvo_nome}.docx")
