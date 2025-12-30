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

# --- 3. FUNÇÕES DE FORMATAÇÃO (ESTRUTURA ABNT/PCPE) ---

def formatar_texto(run, tamanho=11, negrito=False, italico=False):
    """Aplica fonte Arial e formatação de caractere."""
    run.font.name = 'Arial'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.italic = italico

def configurar_paragrafo(paragrafo, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, espaco_depois=0, entrelinhas=1.0, recuo=0):
    """Configura o layout do parágrafo (espaçamentos e alinhamentos)."""
    p_fmt = paragrafo.paragraph_format
    p_fmt.alignment = alinhamento
    p_fmt.space_after = Pt(espaco_depois)
    p_fmt.line_spacing = entrelinhas
    if recuo > 0:
        p_fmt.first_line_indent = Cm(recuo)

# --- 4. CONFIGURAÇÃO DO CABEÇALHO (AJUSTADO PARA FICAR IDÊNTICO) ---
def criar_cabecalho_rodape(doc):
    section = doc.sections[0]
    
    # Margens do Modelo Alex
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.5)
    section.header_distance = Inches(0.2)
    section.footer_distance = Inches(0.2)

    # --- CABEÇALHO ---
    header = section.header
    # Tabela 1x2 para Logo e Texto
    table = header.add_table(rows=1, cols=2, width=Inches(6.8))
    table.autofit = False
    table.columns[0].width = Inches(1.1) # Espaço do Logo
    table.columns[1].width = Inches(5.7) # Espaço do Texto

    # Célula 1: Logo
    try:
        cell_logo = table.cell(0, 0)
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_logo = p_logo.add_run()
        run_logo.add_picture('logo_pc.png', width=Inches(0.95))
    except:
        table.cell(0, 0).text = "[LOGO]"

    # Célula 2: Texto Institucional (AUMENTEI A FONTE AQUI)
    cell_text = table.cell(0, 1)
    p_text = cell_text.paragraphs[0]
    p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Título Principal (Maior)
    r1 = p_text.add_run("POLÍCIA CIVIL DE PERNAMBUCO\n")
    formatar_texto(r1, tamanho=14, negrito=True) # Aumentado para 14
    
    # Subtítulos
    r2 = p_text.add_run("DINTER 1 - 16ª DESEC\n")
    formatar_texto(r2, tamanho=11, negrito=True) # Aumentado para 11
    
    r3 = p_text.add_run("Delegacia de Polícia da 116ª Circunscrição - Surubim")
    formatar_texto(r3, tamanho=11, negrito=True) # Aumentado para 11

    # --- RODAPÉ ---
    footer = section.footer
    p_foot = footer.paragraphs[0]
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_foot.add_run("Av. São Sebastião - Surubim - PE | Fone: (81) 3624-1974\nE-mail: dp116circ.surubim@policiacivil.pe.gov.br")
    formatar_texto(r_foot, tamanho=9) # Tamanho 9 para o rodapé

# --- 5. INTERFACE (CAMPOS VAZIOS / GENÉRICOS) ---
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# Barra Lateral
with st.sidebar:
    st.header("1. Dados do Documento")
    # Título editável (Começa genérico, mas preenchido com padrão útil)
    titulo_doc = st.text_input("Título do Relatório:", value="RELATÓRIO DE INVESTIGAÇÃO")
    
    st.markdown("---")
    opj = st.text_input("OPJ:", placeholder="Ex: INTERCEPTUM")
    processo = st.text_input("Processo / BO:", placeholder="Ex: 0002343-02...")
    natureza = st.text_input("Natureza:", placeholder="Ex: Homicídio, Tráfico...")
    
    c1, c2 = st.columns(2)
    data_input = c1.text_input("Data:", placeholder="DD de mês de AAAA")
    hora_input = c2.text_input("Hora:", placeholder="00h00")
    
    local = st.text_input("Local:", placeholder="Endereço da diligência...")

# Título do App
st.title("🚓 Gerador PCPE (Multi-Uso)")

# Abas
tab1, tab2, tab3, tab4 = st.tabs(["👤 Envolvidos", "📝 Relato", "📸 Fotos", "👮 Equipe"])

with tab1:
    st.subheader("Quem são os envolvidos?")
    c_a, c_b = st.columns(2)
    with c_a:
        st.markdown("**Suspeito / Alvo**")
        alvo = st.text_input("Nome do Alvo:")
        cpf_rg = st.text_input("Docs (CPF/RG):")
        nasc = st.text_input("Nascimento / Idade:")
    with c_b:
        st.markdown("**Outros**")
        vitima = st.text_input("Nome da Vítima:")
        advogado = st.text_input("Advogado:")
        testemunha = st.text_input("Testemunha:")

fotos_carregadas = []
with tab3:
    st.info("Suba as fotos e use os códigos [FOTO1], [FOTO2] no texto.")
    fotos_carregadas = st.file_uploader("Imagens", accept_multiple_files=True)
    if fotos_carregadas:
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i%5]:
                st.image(f, width=80)
                st.code(f"[FOTO{i+1}]")

with tab2:
    st.subheader("Corpo do Relatório")
    texto_relato = st.text_area("Descreva os fatos detalhadamente:", height=450, 
        placeholder="Digite aqui o histórico da ocorrência...\n\nUse [FOTO1] para inserir imagens entre parágrafos.")

with tab4:
    st.subheader("Quem assina?")
    agentes = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        n = c1.text_input(f"Nome Agente {i+1}", key=f"n{i}")
        c = c2.text_input(f"Cargo {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes.append((n, c))
    st.button("➕ Adicionar Assinatura", on_click=add_agente)
    st.button("➖ Remover", on_click=remove_agente)

# --- 6. GERAÇÃO ---
st.markdown("---")
if st.button("GERAR RELATÓRIO OFICIAL", type="primary"):
    doc = Document()
    
    # 1. Configura Cabeçalho e Rodapé (Repetem em todas as páginas)
    criar_cabecalho_rodape(doc)
    
    # 2. Título Centralizado
    p_tit = doc.add_paragraph()
    r_tit = p_tit.add_run(titulo_doc.upper()) # Força Maiúscula
    formatar_texto(r_tit, tamanho=12, negrito=True)
    configurar_paragrafo(p_tit, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 3. Bloco de Dados Iniciais (Dinâmico: só mostra o que foi preenchido)
    def add_dado(chave, valor):
        if valor: # Só cria a linha se tiver texto
            p = doc.add_paragraph()
            r_k = p.add_run(f"{chave}: ")
            formatar_texto(r_k, negrito=True)
            r_v = p.add_run(str(valor))
            formatar_texto(r_v, negrito=False)
            configurar_paragrafo(p, espaco_depois=0) # Sem espaço extra, linha colada

    add_dado("NATUREZA", natureza)
    add_dado("OPERAÇÃO (OPJ)", f"\"{opj}\"" if opj else None)
    add_dado("PROCESSO/BO", processo)
    
    # Data e Hora na mesma linha ou separadas
    if data_input and hora_input:
        add_dado("DATA/HORA", f"{data_input} às {hora_input}")
    elif data_input:
        add_dado("DATA", data_input)
        
    add_dado("LOCAL", local)
    
    doc.add_paragraph() # Espaço de respiro

    # 4. Seção Envolvidos (Genérica)
    # Verifica se existe algum dado de envolvido para criar o título
    if any([alvo, vitima, advogado, testemunha]):
        p_sec1 = doc.add_paragraph()
        r_sec1 = p_sec1.add_run("DOS ENVOLVIDOS")
        formatar_texto(r_sec1, negrito=True)
        configurar_paragrafo(p_sec1, espaco_depois=6)

        if alvo:
            txt_alvo = alvo
            if cpf_rg: txt_alvo += f" | {cpf_rg}"
            add_dado("ALVO/INVESTIGADO", txt_alvo)
            if nasc: add_dado("NASCIMENTO", nasc)
        
        add_dado("VÍTIMA", vitima)
        add_dado("ADVOGADO", advogado)
        add_dado("TESTEMUNHA", testemunha)
        
        doc.add_paragraph()

    # 5. Seção Relato
    p_sec2 = doc.add_paragraph()
    r_sec2 = p_sec2.add_run("DO RELATO / DILIGÊNCIA")
    formatar_texto(r_sec2, negrito=True)
    configurar_paragrafo(p_sec2, espaco_depois=6)

    # 6. Processamento Inteligente do Texto + Fotos
    if texto_relato:
        partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
        
        for parte in partes:
            if parte.isdigit():
                # É código de foto
                idx = int(parte) - 1
                if 0 <= idx < len(fotos_carregadas):
                    f = fotos_carregadas[idx]
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_img = p_img.add_run()
                    run_img.add_picture(f, width=Inches(5.5)) # Largura padrão foto
                    
                    p_leg = doc.add_paragraph()
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_leg = p_leg.add_run(f"Figura {idx+1}")
                    formatar_texto(r_leg, tamanho=9)
                    configurar_paragrafo(p_leg, espaco_depois=12)
            else:
                # É texto normal -> Aplicar formatação de parágrafo correta
                paragrafos_texto = parte.split('\n')
                for par in paragrafos_texto:
                    if par.strip():
                        p = doc.add_paragraph(par)
                        # Formatação: Justificado, 1.5 linhas, Recuo 1.25cm
                        configurar_paragrafo(p, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)
                        for run in p.runs:
                            formatar_texto(run, tamanho=11)

    # 7. Assinaturas
    doc.add_paragraph(); doc.add_paragraph()
    for nome, cargo in agentes:
        if nome:
            doc.add_paragraph() # Espaço extra
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(f"__________________________________________\n{nome}\n{cargo}")
            formatar_texto(r, tamanho=11)

    # Download
    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("⬇️ BAIXAR RELATÓRIO", bio.getvalue(), "Relatorio_Oficial.docx", type="primary")
