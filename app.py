import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# Função para garantir a fonte Arial (Idêntica ao Original)
def estilo(run, tamanho=11, negrito=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito

st.set_page_config(page_title="Gerador PCPE - Modelo Alex", layout="centered")
st.title("🚓 Gerador de Relatório (Modelo Alex sem Rodapé)")

# --- FORMULÁRIO DE ENTRADA ---
with st.form("form_alex"):
    st.subheader("1. Dados do Cabeçalho")
    col1, col2 = st.columns(2)
    with col1:
        opj = st.text_input("OPJ:", "INTERCEPTUM")
        processo = st.text_input("Processo nº:", "0002343-02.2025.8.17.3410")
    with col2:
        data = st.text_input("Data:", "22 de dezembro de 2025")
        hora = st.text_input("Hora:", "14h23")
    
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

    st.subheader("2. Dados do Alvo e Envolvidos")
    alvo_nome = st.text_input("Nome do Alvo:", "ALEX DO CARMO CORREIA")
    alvo_docs = st.text_input("CPF / RG:", "CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
    nascimento = st.text_input("Data de Nascimento:", "15/04/2004")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")

    st.subheader("3. Corpo do Relatório")
    st.info("O texto abaixo será formatado automaticamente. Pode colar seu texto com parágrafos.")
    texto_diligencia = st.text_area("Descrição da Diligência:", height=300, 
        value="Em cumprimento à ordem judicial expedida pela Vara Criminal competente, as equipes deslocaram-se ao endereço supracitado...")

    st.subheader("4. Fotos e Assinatura")
    fotos = st.file_uploader("Anexar Fotos", accept_multiple_files=True)
    responsavel = st.text_input("Nome do Responsável:", "Rafael de Albuquerque Campos")
    cargo = st.text_input("Cargo/Matrícula:", "Investigador de Polícia")

    gerar = st.form_submit_button("GERAR DOCUMENTO")

if gerar:
    doc = Document()
    
    # 1. CONFIGURAÇÃO DE MARGENS (Iguais ao arquivo do Alex)
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # 2. CABEÇALHO (Logo + Texto) - Tabela Invisível
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.1) # Largura para o Logo
    table.columns[1].width = Inches(5.5) # Largura para o Texto
    
    # Logo
    try:
        cell_logo = table.cell(0, 0)
        run_logo = cell_logo.paragraphs[0].add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.9))
    except:
        table.cell(0, 0).text = "[LOGO]"

    # Texto Institucional
    cell_text = table.cell(0, 1)
    p_text = cell_text.paragraphs[0]
    run_text = p_text.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    estilo(run_text, tamanho=10, negrito=True)
    
    doc.add_paragraph() # Espaço

    # 3. TÍTULO DO RELATÓRIO
    p_tit = doc.add_paragraph()
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_tit = p_tit.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    estilo(run_tit, tamanho=12, negrito=True)
    
    doc.add_paragraph() # Espaço

    # 4. METADADOS (OPJ, Processo, etc)
    def add_meta(label, valor):
        p = doc.add_paragraph()
        r1 = p.add_run(f"{label}: ")
        estilo(r1, negrito=True)
        r2 = p.add_run(valor)
        estilo(r2)
        p.paragraph_format.space_after = Pt(2)

    add_meta("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_meta("PROCESSO nº", processo)
    add_meta("DATA", data)
    add_meta("HORA", hora)
    add_meta("LOCAL", local)

    doc.add_paragraph()

    # 5. SEÇÃO DO ALVO (Alinhamento Específico)
    p_sect1 = doc.add_paragraph()
    estilo(p_sect1.add_run("DO ALVO E TESTEMUNHAS"), negrito=True)

    # Linha do Nome + Docs
    p_alvo = doc.add_paragraph()
    estilo(p_alvo.add_run("ALVO: "), negrito=True)
    estilo(p_alvo.add_run(f"{alvo_nome} | {alvo_docs}"))
    
    # Linha Nascimento
    p_nasc = doc.add_paragraph()
    estilo(p_nasc.add_run("Nascimento: "), negrito=True)
    estilo(p_nasc.add_run(nascimento))
    
    # Linha Advogado
    p_adv = doc.add_paragraph()
    estilo(p_adv.add_run("ADVOGADO: "), negrito=True)
    estilo(p_adv.add_run(advogado))

    # Linha Testemunha
    p_test = doc.add_paragraph()
    estilo(p_test.add_run("TESTEMUNHA: "), negrito=True)
    estilo(p_test.add_run(testemunha))

    doc.add_paragraph()

    # 6. SEÇÃO DILIGÊNCIA (Texto Justificado)
    p_sect2 = doc.add_paragraph()
    estilo(p_sect2.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO"), negrito=True)

    # Divide o texto em parágrafos para não embolar
    paragrafos = texto_diligencia.split('\n')
    for par in paragrafos:
        if par.strip():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run_p = p.add_run(par)
            estilo(run_p, 11)
            p.paragraph_format.space_after = Pt(6)

    # 7. FOTOS (Centralizadas)
    if fotos:
        for f in fotos:
            doc.add_page_break()
            # Imagem
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(f, width=Inches(5.5))
            
            # Legenda
            p_leg = doc.add_paragraph()
            p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
            estilo(p_leg.add_run(f"Registro Fotográfico: {f.name}"), 9)

    # 8. ASSINATURA (Sem Rodapé)
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph() # Espaços para assinar
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sig = p_sig.add_run(f"__________________________________________\n{responsavel}\n{cargo}")
    estilo(run_sig, 11)

    # NOTA: O código NÃO adiciona nada ao rodapé (section.footer), então ele ficará vazio.

    # GERAÇÃO DO ARQUIVO
    bio = io.BytesIO()
    doc.save(bio)
    st.success("✅ Relatório gerado com sucesso (Formato Alex - Sem Rodapé)")
    st.download_button("⬇️ Baixar DOCX", bio.getvalue(), "Relatorio_Final_Alex.docx")
