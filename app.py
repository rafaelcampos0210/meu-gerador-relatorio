import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# --- FUNÇÃO DE ESTILO (Segredo da Formatação) ---
def estilo_fiel(run, tamanho=11, negrito=False, italico=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.italic = italico

st.set_page_config(page_title="Gerador PCPE Pro", layout="centered")
st.title("🚓 Gerador de Relatório - Réplica Exata")

# --- FORMULÁRIO ---
with st.form("main_form"):
    st.subheader("1. Cabeçalho e Processo")
    col_a, col_b = st.columns(2)
    with col_a:
        opj = st.text_input("OPJ:", "INTERCEPTUM")
        processo = st.text_input("Processo nº:", "0002343-02.2025.8.17.3410")
    with col_b:
        data_doc = st.text_input("Data:", "22 de dezembro de 2025")
        local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

    st.subheader("2. Dados do Alvo")
    alvo_nome = st.text_input("Nome do Alvo:", "ALEX DO CARMO CORREIA")
    alvo_qualificacao = st.text_input("Docs (CPF/RG):", "CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
    nascimento = st.text_input("Data de Nascimento:", "15/04/2004")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")

    st.subheader("3. Corpo do Relatório")
    relato = st.text_area("Texto da Diligência:", height=300)
    
    st.subheader("4. Fotos")
    fotos = st.file_uploader("Evidências", accept_multiple_files=True)
    
    submit = st.form_submit_button("GERAR DOCUMENTO ORIGINAL")

if submit:
    doc = Document()
    
    # 1. MARGENS EXATAS (0.5 polegadas = 1.27cm)
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # 2. CABEÇALHO (Logo + Texto lado a lado)
    # Tabela 1x2 para travar o layout
    table_head = doc.add_table(rows=1, cols=2)
    table_head.autofit = False 
    table_head.columns[0].width = Inches(1.1)
    table_head.columns[1].width = Inches(5.5)

    # Célula do Logo
    cell_logo = table_head.cell(0, 0)
    try:
        p_logo = cell_logo.paragraphs[0]
        run_logo = p_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.85))
    except:
        cell_logo.text = " [LOGO] "

    # Célula do Texto Institucional
    cell_text = table_head.cell(0, 1)
    p_text = cell_text.paragraphs[0]
    run_text = p_text.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    estilo_fiel(run_text, tamanho=10, negrito=True)
    p_text.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Espaço após cabeçalho
    doc.add_paragraph() 

    # 3. TÍTULO
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    estilo_fiel(run_title, tamanho=12, negrito=True)
    
    # 4. BLOCO DE DADOS (OPJ, Processo, etc.)
    # Usando parágrafos com TABULAÇÃO MANUAL para alinhar
    def add_line(label, content):
        p = doc.add_paragraph()
        r1 = p.add_run(f"{label}: ")
        estilo_fiel(r1, negrito=True)
        r2 = p.add_run(content)
        estilo_fiel(r2)
        p.paragraph_format.space_after = Pt(2) # Espaço pequeno entre linhas

    add_line("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_line("PROCESSO nº", processo)
    add_line("DATA", data_doc)
    add_line("LOCAL", local)

    doc.add_paragraph() # Espaço

    # 5. SEÇÃO ALVO (Alinhamento Crítico)
    p_sect1 = doc.add_paragraph()
    r_sect1 = p_sect1.add_run("DO ALVO E TESTEMUNHAS")
    estilo_fiel(r_sect1, negrito=True)

    # Alvo
    p_alvo = doc.add_paragraph()
    r_a1 = p_alvo.add_run("ALVO: ")
    estilo_fiel(r_a1, negrito=True)
    r_a2 = p_alvo.add_run(f"{alvo_nome} | {alvo_qualificacao}")
    estilo_fiel(r_a2)
    
    # Nascimento (Linha separada para não embolar)
    p_nasc = doc.add_paragraph()
    r_n1 = p_nasc.add_run("Nascimento: ")
    estilo_fiel(r_n1, negrito=True)
    r_n2 = p_nasc.add_run(nascimento)
    estilo_fiel(r_n2)

    # Advogado
    p_adv = doc.add_paragraph()
    r_ad1 = p_adv.add_run("ADVOGADO: ")
    estilo_fiel(r_ad1, negrito=True)
    r_ad2 = p_adv.add_run(advogado)
    estilo_fiel(r_ad2)

    # Testemunha
    p_test = doc.add_paragraph()
    r_t1 = p_test.add_run("TESTEMUNHA: ")
    estilo_fiel(r_t1, negrito=True)
    r_t2 = p_test.add_run(testemunha)
    estilo_fiel(r_t2)

    doc.add_paragraph()

    # 6. SEÇÃO DILIGÊNCIA
    p_sect2 = doc.add_paragraph()
    r_sect2 = p_sect2.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    estilo_fiel(r_sect2, negrito=True)

    p_relato = doc.add_paragraph()
    p_relato.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r_relato = p_relato.add_run(relato)
    estilo_fiel(r_relato)

    # 7. FOTOS
    if fotos:
        for foto in fotos:
            doc.add_page_break()
            # Foto Centralizada
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(foto, width=Inches(5.8)) # Largura máxima da margem
            
            # Legenda Centralizada
            p_leg = doc.add_paragraph()
            p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_leg = p_leg.add_run(f"Registro Fotográfico: {foto.name}")
            estilo_fiel(r_leg, tamanho=9)

    # 8. RODAPÉ FIXO
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 36241974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    estilo_fiel(r_foot, tamanho=8)

    # GERAÇÃO
    bio = io.BytesIO()
    doc.save(bio)
    
    st.success("✅ Relatório gerado com sucesso!")
    st.download_button("⬇️ Baixar DOCX Final", bio.getvalue(), "Relatorio_PCPE_Final.docx")
