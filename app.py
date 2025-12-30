import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
import io
import re

# --- 1. CONFIGURAÇÃO VISUAL ---
st.set_page_config(page_title="Gerador PCPE Oficial", layout="wide", page_icon="🚓")

st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: 500; color: #000;}
    .stTextArea textarea {font-family: 'Arial'; font-size: 14px;}
    /* Estilo para a galeria de fotos */
    .foto-container {
        background-color: white;
        padding: 10px;
        border-radius: 8px;
        border: 1px solid #ddd;
        margin-bottom: 10px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. FUNÇÕES DE FORMATAÇÃO ---
def formatar_texto(run, tamanho=11, negrito=False, italico=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.italic = italico
    run.font.color.rgb = RGBColor(0, 0, 0)

def configurar_paragrafo(paragrafo, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espaco_depois=0, entrelinhas=1.0, recuo=0):
    p_fmt = paragrafo.paragraph_format
    p_fmt.alignment = alinhamento
    p_fmt.space_after = Pt(espaco_depois)
    p_fmt.line_spacing = entrelinhas
    if recuo > 0: p_fmt.first_line_indent = Cm(recuo)

# --- 3. CABEÇALHO PERFEITO (3 COLUNAS BALANCEADAS) ---
def criar_cabecalho_rodape(doc):
    section = doc.sections[0]
    
    # Margens (Ajustadas para caber o cabeçalho largo)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8) # ~2.0 cm
    section.right_margin = Inches(0.5) # ~1.27 cm
    section.header_distance = Inches(0.2)
    section.footer_distance = Inches(0.2)

    # --- CABEÇALHO ---
    header = section.header
    
    # Largura total útil da página = 8.5" (folha) - 0.8" (esq) - 0.5" (dir) = 7.2"
    largura_total = 7.2
    largura_lateral = 1.3 # Espaço para o Logo (Esquerda) e Vazio (Direita)
    largura_central = largura_total - (largura_lateral * 2) # O que sobra pro texto (4.6")
    
    # Cria tabela 1x3
    table = header.add_table(rows=1, cols=3, width=Inches(largura_total))
    table.autofit = False
    
    # Define as larguras EXATAS
    table.columns[0].width = Inches(largura_lateral) # Coluna 1 (Logo)
    table.columns[1].width = Inches(largura_central) # Coluna 2 (Texto)
    table.columns[2].width = Inches(largura_lateral) # Coluna 3 (Equilíbrio)

    # COLUNA 1: LOGO (Alinhado à Esquerda)
    try:
        cell_logo = table.cell(0, 0)
        cell_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT # Encostado na margem
        run_logo = p_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(1.0))
    except:
        table.cell(0, 0).text = "[LOGO]"

    # COLUNA 2: TEXTO (Centralizado na Célula -> Centralizado na Página)
    cell_text = table.cell(0, 1)
    cell_text.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_text._element.clear_content()

    def add_line(texto, tamanho):
        p = cell_text.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0 # Simples
        r = p.add_run(texto)
        formatar_texto(r, tamanho=tamanho, negrito=True)

    add_line("POLÍCIA CIVIL DE PERNAMBUCO", 14)
    add_line("DINTER 1 - 16ª DESEC", 11)
    add_line("Delegacia de Polícia da 116ª Circunscrição - Surubim", 11)

    # COLUNA 3: VAZIA (Essencial para o equilíbrio)
    # Ela "empurra" o texto para o centro exato.

    # --- RODAPÉ ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 3624-1974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    formatar_texto(r_foot, tamanho=9)

# --- 4. INTERFACE ---
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

with st.sidebar:
    st.header("1. Cabeçalho")
    titulo_doc = st.text_input("Título:", value="RELATÓRIO DE INVESTIGAÇÃO")
    st.markdown("---")
    opj = st.text_input("OPJ:", placeholder="Ex: INTERCEPTUM")
    processo = st.text_input("Processo:", placeholder="0002343...")
    natureza = st.text_input("Natureza:", placeholder="Homicídio...")
    c1, c2 = st.columns(2)
    data_input = c1.text_input("Data:", placeholder="DD/MM/AAAA")
    hora_input = c2.text_input("Hora:", placeholder="HH:MM")
    local = st.text_input("Local:", placeholder="Endereço...")

st.title("🚓 Gerador PCPE (Layout Fixo)")

# Variável global de fotos
fotos_carregadas = []

# ABAS
tab1, tab2, tab3 = st.tabs(["📝 Relato e Fotos (Integrado)", "👤 Envolvidos", "👮 Equipe"])

with tab1:
    col_upload, col_texto = st.columns([1, 2])
    
    with col_upload:
        st.info("1. Selecione as Fotos")
        fotos_carregadas = st.file_uploader("Upload", accept_multiple_files=True, label_visibility="collapsed")
        
        if fotos_carregadas:
            st.markdown("---")
            st.write("📋 **Galeria de Tags**")
            st.caption("Clique no código para copiar")
            
            # Galeria Vertical para facilitar
            for i, f in enumerate(fotos_carregadas):
                with st.container():
                    c_img, c_code = st.columns([1, 2])
                    c_img.image(f, width=60)
                    # O st.code cria um botão de copiar nativo
                    c_code.code(f"[FOTO{i+1}]", language="html")

    with col_texto:
        st.subheader("2. Redação do Relatório")
        st.markdown("Escreva o texto e cole os códigos `[FOTO...]` onde a imagem deve aparecer.")
        texto_relato = st.text_area("Corpo do Texto:", height=600, 
                                   placeholder="Ex: A equipe chegou ao local... \n\n[FOTO1]\n\nFoi encontrado...")

with tab2:
    st.subheader("Envolvidos")
    c_a, c_b = st.columns(2)
    with c_a:
        alvo = st.text_input("Nome Alvo:")
        cpf_rg = st.text_input("Docs (CPF/RG):")
        nasc = st.text_input("Nascimento:")
    with c_b:
        vitima = st.text_input("Nome Vítima:")
        advogado = st.text_input("Advogado:")
        testemunha = st.text_input("Testemunha:")

with tab3:
    st.subheader("Assinaturas")
    agentes = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        n = c1.text_input(f"Nome {i+1}", key=f"n{i}")
        c = c2.text_input(f"Cargo {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes.append((n, c))
    st.button("➕ Adicionar", on_click=add_agente)
    st.button("➖ Remover", on_click=remove_agente)

# --- 5. GERAÇÃO ---
st.markdown("---")
if st.button("GERAR RELATÓRIO FINAL", type="primary"):
    doc = Document()
    
    # 1. Cabeçalho
    criar_cabecalho_rodape(doc)
    
    # 2. Título
    p_tit = doc.add_paragraph()
    r_tit = p_tit.add_run(titulo_doc.upper())
    formatar_texto(r_tit, tamanho=12, negrito=True)
    configurar_paragrafo(p_tit, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 3. Dados
    def add_dado(chave, valor):
        if valor:
            p = doc.add_paragraph()
            r_k = p.add_run(f"{chave}: ")
            formatar_texto(r_k, negrito=True)
            r_v = p.add_run(str(valor))
            formatar_texto(r_v, negrito=False)
            configurar_paragrafo(p, espaco_depois=0)

    add_dado("NATUREZA", natureza)
    add_dado("OPERAÇÃO (OPJ)", f"\"{opj}\"" if opj else None)
    add_dado("PROCESSO/BO", processo)
    if data_input and hora_input:
        add_dado("DATA/HORA", f"{data_input} às {hora_input}")
    elif data_input:
        add_dado("DATA", data_input)
    add_dado("LOCAL", local)
    
    doc.add_paragraph()

    # 4. Envolvidos
    if any([alvo, vitima, advogado, testemunha]):
        p_sec1 = doc.add_paragraph()
        r_sec1 = p_sec1.add_run("DOS ENVOLVIDOS")
        formatar_texto(r_sec1, negrito=True)
        configurar_paragrafo(p_sec1, espaco_depois=6)

        if alvo:
            txt = alvo
            if cpf_rg: txt += f" | {cpf_rg}"
            add_dado("ALVO/INVESTIGADO", txt)
            if nasc: add_dado("NASCIMENTO", nasc)
        add_dado("VÍTIMA", vitima)
        add_dado("ADVOGADO", advogado)
        add_dado("TESTEMUNHA", testemunha)
        doc.add_paragraph()

    # 5. Relato
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run("DO RELATO / DILIGÊNCIA")
    formatar_texto(r_sec2, negrito=True)
    configurar_paragrafo(p_sec2, espaco_depois=6)

    # 6. Processamento
    if texto_relato:
        partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
        for parte in partes:
            if parte.isdigit():
                idx = int(parte) - 1
                if 0 <= idx < len(fotos_carregadas):
                    f = fotos_carregadas[idx]
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_img = p_img.add_run()
                    run_img.add_picture(f, width=Inches(5.5))
                    p_leg = doc.add_paragraph()
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_leg = p_leg.add_run(f"Figura {idx+1}")
                    formatar_texto(r_leg, tamanho=9)
                    configurar_paragrafo(p_leg, espaco_depois=12)
            else:
                for par in parte.split('\n'):
                    if par.strip():
                        p = doc.add_paragraph(par)
                        configurar_paragrafo(p, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)
                        for run in p.runs: formatar_texto(run, tamanho=11)

    # 7. Assinaturas
    doc.add_paragraph(); doc.add_paragraph()
    for nome, cargo in agentes:
        if nome:
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"__________________________________________\n{nome}\n{cargo}")
            formatar_texto(r, tamanho=11)

    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("⬇️ BAIXAR DOCX", bio.getvalue(), "Relatorio_PCPE_Final.docx", type="primary")
