import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador PCPE Oficial", layout="wide", page_icon="🚓")

# --- 2. ESTILO VISUAL (CSS) ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: 500; color: #000;}
    .stTextArea textarea {font-family: 'Arial'; font-size: 14px;}
    .tag-foto {
        background-color: #e3f2fd; border: 1px solid #1565c0; color: #1565c0; 
        padding: 2px 8px; border-radius: 4px; font-weight: bold; font-family: monospace;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. FUNÇÕES DE FORMATAÇÃO ---

def formatar_texto(run, tamanho=11, negrito=False, italico=False):
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.italic = italico

def configurar_paragrafo(paragrafo, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espaco_depois=0, entrelinhas=1.0, recuo=0):
    p_fmt = paragrafo.paragraph_format
    p_fmt.alignment = alinhamento
    p_fmt.space_after = Pt(espaco_depois)
    p_fmt.line_spacing = entrelinhas
    if recuo > 0:
        p_fmt.first_line_indent = Cm(recuo)

# --- 4. CONFIGURAÇÃO DO CABEÇALHO (TÉCNICA DE 3 COLUNAS) ---
def criar_cabecalho_rodape(doc):
    section = doc.sections[0]
    
    # Margens
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.5)
    section.header_distance = Inches(0.2)
    section.footer_distance = Inches(0.2)

    # --- CABEÇALHO ---
    header = section.header
    
    # TRUQUE: Tabela de 3 colunas para garantir que o texto fique no CENTRO DA PÁGINA
    # Col 1 (Logo) | Col 2 (Texto) | Col 3 (Vazia para equilíbrio)
    # Largura Total ~ 7.0 polegadas (Margem util)
    table = header.add_table(rows=1, cols=3, width=Inches(7.0))
    table.autofit = False
    
    # Define larguras exatas para balancear
    largura_lateral = Inches(1.2)
    largura_central = Inches(4.6)
    
    table.columns[0].width = largura_lateral # Esquerda (Logo)
    table.columns[1].width = largura_central # Centro (Texto)
    table.columns[2].width = largura_lateral # Direita (Vazia) - O segredo está aqui!

    # 1. Coluna Esquerda: LOGO
    try:
        cell_logo = table.cell(0, 0)
        cell_logo.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT # Logo alinhado à esquerda
        run_logo = p_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(1.0))
    except:
        table.cell(0, 0).text = "[LOGO]"

    # 2. Coluna Central: TEXTO (Centralizado)
    cell_text = table.cell(0, 1)
    cell_text.vertical_alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_text = cell_text.paragraphs[0]
    p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Título Principal
    r1 = p_text.add_run("POLÍCIA CIVIL DE PERNAMBUCO\n")
    formatar_texto(r1, tamanho=16, negrito=True)
    
    # Subtítulos
    r2 = p_text.add_run("DINTER 1 - 16ª DESEC\n")
    formatar_texto(r2, tamanho=12, negrito=True)
    
    r3 = p_text.add_run("Delegacia de Polícia da 116ª Circunscrição - Surubim")
    formatar_texto(r3, tamanho=12, negrito=True)
    
    # 3. Coluna Direita: VAZIA (Não fazemos nada, ela existe só para empurrar o centro para o meio)

    # --- RODAPÉ ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 3624-1974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    formatar_texto(r_foot, tamanho=9)

# --- 5. INTERFACE DO USUÁRIO ---
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

with st.sidebar:
    st.header("1. Cabeçalho")
    titulo_doc = st.text_input("Título do Relatório:", value="RELATÓRIO DE INVESTIGAÇÃO")
    
    st.markdown("---")
    opj = st.text_input("OPJ:", placeholder="Ex: INTERCEPTUM")
    processo = st.text_input("Processo / BO:", placeholder="Ex: 0002343...")
    natureza = st.text_input("Natureza:", placeholder="Ex: Homicídio...")
    
    c1, c2 = st.columns(2)
    data_input = c1.text_input("Data:", placeholder="DD/MM/AAAA")
    hora_input = c2.text_input("Hora:", placeholder="HH:MM")
    
    local = st.text_input("Local:", placeholder="Endereço...")

st.title("🚓 Gerador PCPE (Formato Final Ajustado)")

tab1, tab2, tab3, tab4 = st.tabs(["👤 Envolvidos", "📝 Relato", "📸 Fotos", "👮 Equipe"])

with tab1:
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

fotos_carregadas = []
with tab3:
    st.info("Use [FOTO1], [FOTO2] no texto.")
    fotos_carregadas = st.file_uploader("Imagens", accept_multiple_files=True)
    if fotos_carregadas:
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i%5]:
                st.image(f, width=80)
                st.code(f"[FOTO{i+1}]")

with tab2:
    st.subheader("Relato Policial")
    texto_relato = st.text_area("Descreva os fatos:", height=450, 
        placeholder="Escreva aqui...")

with tab4:
    st.subheader("Assinaturas")
    agentes = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        n = c1.text_input(f"Nome {i+1}", key=f"n{i}")
        c = c2.text_input(f"Cargo {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes.append((n, c))
    st.button("➕ Adicionar", on_click=add_agente)
    st.button("➖ Remover", on_click=remove_agente)

# --- 6. GERAÇÃO ---
st.markdown("---")
if st.button("GERAR RELATÓRIO CORRIGIDO", type="primary"):
    doc = Document()
    
    # 1. Cabeçalho Corrigido (3 Colunas)
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
    st.download_button("⬇️ BAIXAR DOCX CORRIGIDO", bio.getvalue(), "Relatorio_PCPE_Final.docx", type="primary")
