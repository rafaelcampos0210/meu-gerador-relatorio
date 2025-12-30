import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# Função para forçar fonte Arial (Padrão PCPE)
def estilo(run, tamanho=11, negrito=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE Pro", layout="centered")
st.title("🚓 Gerador de Relatório (Corrigido)")

with st.form("main"):
    col1, col2 = st.columns(2)
    with col1:
        opj = st.text_input("OPJ:", "INTERCEPTUM")
        processo = st.text_input("Processo:", "0002343-02.2025.8.17.3410")
    with col2:
        data = st.text_input("Data:", "22 de dezembro de 2025")
        local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural...")

    st.markdown("---")
    st.caption("Dados do Alvo")
    alvo = st.text_input("Alvo Completo:", "ALEX DO CARMO CORREIA | CPF: ...")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo...")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene...")
    
    st.markdown("---")
    st.caption("Texto do Relatório")
    # CORREÇÃO AQUI: help= (com sinal de igual)
    relato = st.text_area("Descrição:", height=400, help="O texto manterá os parágrafos originais.")
    
    fotos = st.file_uploader("Fotos", accept_multiple_files=True)
    btn = st.form_submit_button("GERAR RELATÓRIO")

if btn:
    doc = Document()
    
    # 1. Configurar Margens
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)

    # 2. Cabeçalho com Logo (Tabela)
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.1)
    
    # Logo
    try:
        run_img = table.cell(0,0).paragraphs[0].add_run()
        run_img.add_picture('logo_pc.png', width=Inches(0.85))
    except:
        table.cell(0,0).text = "[LOGO]"
        
    # Texto Cabeçalho
    p_head = table.cell(0,1).paragraphs[0]
    run_head = p_head.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1 - 16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    estilo(run_head, 10, True)

    doc.add_paragraph() # Espaço

    # 3. Título
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    estilo(run, 12, True)

    # 4. Dados Técnicos
    def add_dado(label, valor):
        p = doc.add_paragraph()
        r1 = p.add_run(f"{label}: ")
        estilo(r1, negrito=True)
        r2 = p.add_run(valor)
        estilo(r2)
        p.paragraph_format.space_after = Pt(2)

    add_dado("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_dado("PROCESSO nº", processo)
    add_dado("DATA", data)
    add_dado("LOCAL", local)

    doc.add_paragraph()

    # 5. Seção Alvo
    p = doc.add_paragraph()
    estilo(p.add_run("DO ALVO E TESTEMUNHAS"), negrito=True)
    
    add_dado("ALVO", alvo)
    add_dado("ADVOGADO", advogado)
    add_dado("TESTEMUNHA", testemunha)

    doc.add_paragraph()

    # 6. SEÇÃO DILIGÊNCIA (Quebra de parágrafos correta)
    p = doc.add_paragraph()
    estilo(p.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO"), negrito=True)

    # Divide o texto onde você deu Enter
    paragrafos = relato.split('\n') 
    
    for paragrafo in paragrafos:
        if paragrafo.strip(): # Só adiciona se tiver texto
            p_novo = doc.add_paragraph()
            p_novo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run_p = p_novo.add_run(paragrafo)
            estilo(run_p, 11)
            p_novo.paragraph_format.space_after = Pt(6)

    # 7. Fotos
    if fotos:
        for f in fotos:
            doc.add_page_break()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(f, width=Inches(5.8))
            
            p_leg = doc.add_paragraph()
            p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
            estilo(p_leg.add_run(f"Registro Fotográfico: {f.name}"), 9)

    # 8. Rodapé
    foot = sec.footer.paragraphs[0]
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    estilo(foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 3624-1974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br"), 8)

    bio = io.BytesIO()
    doc.save(bio)
    st.download_button("⬇️ Baixar Relatório Corrigido", bio.getvalue(), "Relatorio_PCPE_vFinal.docx")
