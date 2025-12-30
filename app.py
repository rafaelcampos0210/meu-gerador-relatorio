import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import io
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador PCPE Oficial", layout="wide", page_icon="🚓")

# --- 2. ESTILO VISUAL DO SITE ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: bold; color: #1f2c56;}
    .stTextArea textarea {font-family: 'Arial'; font-size: 14px;}
    .tag-foto {
        background-color: #e3f2fd; border: 1px solid #1565c0; color: #1565c0; 
        padding: 2px 8px; border-radius: 4px; font-weight: bold; font-family: monospace;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. FUNÇÕES DE FORMATAÇÃO (ESTRUTURA ABNT/PCPE) ---

def formatar_texto(run, tamanho=11, negrito=False, cor_rgb=None):
    """Aplica fonte Arial, tamanho e cor."""
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito
    if cor_rgb:
        run.font.color.rgb = cor_rgb

def configurar_paragrafo(paragrafo, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espaco_depois=0, entrelinhas=1.0, recuo=0):
    """Configura o layout do parágrafo."""
    p_fmt = paragrafo.paragraph_format
    p_fmt.alignment = alinhamento
    p_fmt.space_after = Pt(espaco_depois)
    p_fmt.line_spacing = entrelinhas # 1.0 = Simples, 1.5 = 1,5 linhas
    if recuo > 0:
        p_fmt.first_line_indent = Cm(recuo)

# --- 4. CONFIGURAÇÃO DO CABEÇALHO E RODAPÉ (REPETIR EM TODAS AS PÁGINAS) ---
def criar_cabecalho_rodape(doc):
    section = doc.sections[0]
    
    # --- MARGENS (IDÊNTICAS AO MODELO) ---
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8) # Margem esquerda levemente maior para encadernação
    section.right_margin = Inches(0.5)
    section.header_distance = Inches(0.2)
    section.footer_distance = Inches(0.2)

    # --- CABEÇALHO (HEADER) ---
    header = section.header
    # Cria tabela invisível 1x2 para Logo e Texto
    table = header.add_table(rows=1, cols=2, width=Inches(6.5))
    table.autofit = False
    table.columns[0].width = Inches(1.1) # Coluna do Logo
    table.columns[1].width = Inches(5.4) # Coluna do Texto

    # Célula 1: Logo
    try:
        cell_logo = table.cell(0, 0)
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_logo = p_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.95)) # Tamanho exato do brasão
    except:
        table.cell(0, 0).text = "[LOGO]"

    # Célula 2: Texto Institucional
    cell_text = table.cell(0, 1)
    p_text = cell_text.paragraphs[0]
    p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER # Texto centralizado na célula
    
    # Linha 1
    r1 = p_text.add_run("POLÍCIA CIVIL DE PERNAMBUCO\n")
    formatar_texto(r1, tamanho=12, negrito=True) # Arial 12 Negrito
    # Linha 2
    r2 = p_text.add_run("DINTER 1 - 16ª DESEC\n")
    formatar_texto(r2, tamanho=10, negrito=True)
    # Linha 3
    r3 = p_text.add_run("Delegacia de Polícia da 116ª Circunscrição - Surubim")
    formatar_texto(r3, tamanho=10, negrito=True)

    # --- RODAPÉ (FOOTER) ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Texto padrão do rodapé PCPE (Cor azul escuro ou preto, vou usar preto padrão)
    r_foot = p_foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 3624-1974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    formatar_texto(r_foot, tamanho=8) # Fonte pequena no rodapé

# --- 5. INTERFACE DO USUÁRIO ---
# Gerenciamento de Agentes
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# Barra Lateral
with st.sidebar:
    st.header("1. Cabeçalho do Relatório")
    titulo_doc = st.text_input("Título:", value="RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    
    st.markdown("---")
    opj = st.text_input("OPJ:", value="INTERCEPTUM")
    processo = st.text_input("Processo:", value="0002343-02.2025.8.17.3410")
    
    c1, c2 = st.columns(2)
    data_input = c1.text_input("Data:", "22 de dezembro de 2025")
    hora_input = c2.text_input("Hora:", "14h23")
    
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

# Título do Site
st.title("🚓 Gerador PCPE (Formato Fiel)")

# Abas
tab1, tab2, tab3, tab4 = st.tabs(["👤 Envolvidos", "📝 Relato", "📸 Fotos", "👮 Equipe"])

with tab1:
    st.subheader("Dados dos Envolvidos")
    c_a, c_b = st.columns(2)
    with c_a:
        alvo = st.text_input("Alvo:", "ALEX DO CARMO CORREIA")
        cpf_rg = st.text_input("Docs (CPF/RG):", "CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
        nasc = st.text_input("Nascimento:", "15/04/2004")
    with c_b:
        advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
        testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")

fotos_carregadas = []
with tab3:
    st.info("Faça o upload das fotos e use o código [FOTO1], [FOTO2] no texto.")
    fotos_carregadas = st.file_uploader("Imagens", accept_multiple_files=True)
    if fotos_carregadas:
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i%5]:
                st.image(f, width=80)
                st.code(f"[FOTO{i+1}]")

with tab2:
    st.subheader("Texto da Diligência")
    texto_relato = st.text_area("Descreva os fatos (Use [FOTO1] para inserir imagens):", height=400, 
        placeholder="Em cumprimento à ordem judicial...\n\n[FOTO1]\n\nFoi localizado...")

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

# --- 6. GERAÇÃO DO DOCUMENTO ---
st.markdown("---")
if st.button("GERAR RELATÓRIO IDÊNTICO AO MODELO", type="primary"):
    doc = Document()
    
    # 1. Aplica o Cabeçalho e Rodapé em TODAS as páginas
    criar_cabecalho_rodape(doc)
    
    # 2. Título do Documento
    p_tit = doc.add_paragraph()
    r_tit = p_tit.add_run(titulo_doc)
    formatar_texto(r_tit, tamanho=12, negrito=True)
    configurar_paragrafo(p_tit, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 3. Bloco de Dados (Compacto - Espaçamento Simples)
    def add_dado(chave, valor):
        p = doc.add_paragraph()
        r_k = p.add_run(f"{chave}: ")
        formatar_texto(r_k, negrito=True)
        r_v = p.add_run(valor)
        formatar_texto(r_v, negrito=False)
        # Espaçamento exato do modelo (sem espaço extra entre linhas de dados)
        configurar_paragrafo(p, espaco_depois=0, entrelinhas=1.0)

    add_dado("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_dado("PROCESSO nº", processo)
    add_dado("DATA", data_input)
    if hora_input: add_dado("HORA", hora_input)
    add_dado("LOCAL", local)
    
    doc.add_paragraph() # Espaço vazio

    # 4. Seção Alvo
    p_sec1 = doc.add_paragraph()
    r_sec1 = p_sec1.add_run("DO ALVO E TESTEMUNHAS")
    formatar_texto(r_sec1, negrito=True)
    configurar_paragrafo(p_sec1, espaco_depois=6) # Espaço pequeno após título

    add_dado("ALVO", f"{alvo} | {cpf_rg}")
    add_dado("Nascimento", nasc)
    add_dado("ADVOGADO", advogado)
    add_dado("TESTEMUNHA", testemunha)
    
    doc.add_paragraph()

    # 5. Seção Diligência
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    formatar_texto(r_sec2, negrito=True)
    configurar_paragrafo(p_sec2, espaco_depois=6)

    # 6. Processamento do Texto + Fotos
    partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
    
    for parte in partes:
        if parte.isdigit():
            # É uma foto
            idx = int(parte) - 1
            if 0 <= idx < len(fotos_carregadas):
                f = fotos_carregadas[idx]
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img = p_img.add_run()
                run_img.add_picture(f, width=Inches(5.5))
                
                p_leg = doc.add_paragraph()
                p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_leg = p_leg.add_run(f"Registro Fotográfico: {f.name}")
                formatar_texto(r_leg, tamanho=9)
                configurar_paragrafo(p_leg, espaco_depois=12)
        else:
            # É texto
            paragrafos_texto = parte.split('\n')
            for par in paragrafos_texto:
                if par.strip():
                    p = doc.add_paragraph(par)
                    # Formatação do Texto: Justificado, 1.5 linhas, Recuo 1.25cm
                    configurar_paragrafo(p, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)
                    # Aplica fonte em todo o parágrafo
                    for run in p.runs:
                        formatar_texto(run, tamanho=11)

    # 7. Assinaturas
    doc.add_paragraph(); doc.add_paragraph()
    for nome, cargo in agentes:
        if nome:
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"__________________________________________\n{nome}\n{cargo}")
            formatar_texto(r, tamanho=11)

    # Download
    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("⬇️ BAIXAR RELATÓRIO FIEL", bio.getvalue(), "Relatorio_PCPE_Oficial.docx", type="primary")
