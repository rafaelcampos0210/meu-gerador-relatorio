import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# --- CONFIGURAÇÃO DA PÁGINA (Layout Wide) ---
st.set_page_config(page_title="Relatório Policial Pro", layout="wide", page_icon="🚓")

# --- ESTILO CSS PERSONALIZADO (Para ficar bonito) ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1 {color: #1f2c56;}
    .stButton>button {width: 100%; border-radius: 5px; height: 3em; font-weight: bold;}
    .stTextArea textarea {font-size: 14px;}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE FORMATAÇÃO (Mantém o padrão Oficial) ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    paragrafo.style.font.name = 'Arial'
    paragrafo.style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois)
    p_format.line_spacing = entrelinhas
    if recuo > 0: p_format.first_line_indent = Cm(recuo)
    if alinhamento is not None: paragrafo.alignment = alinhamento
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito

# --- GERENCIAMENTO DE ESTADO (Para adicionar agentes dinamicamente) ---
if 'num_agentes' not in st.session_state:
    st.session_state.num_agentes = 2 # Começa com 2 agentes

def add_agente():
    st.session_state.num_agentes += 1

def remove_agente():
    if st.session_state.num_agentes > 1:
        st.session_state.num_agentes -= 1

# --- BARRA LATERAL (Configurações do Documento) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2502/2502923.png", width=80)
    st.title("Configurações")
    st.markdown("---")
    opj = st.text_input("Nome da OPJ:", "INTERCEPTUM")
    processo = st.text_input("Nº Processo:", "0002343-02.2025.8.17.3410")
    data_doc = st.date_input("Data do Documento:")
    hora_doc = st.time_input("Hora do Documento:")
    local = st.text_input("Local da Diligência:", "Sítio Salvador, nº 360, Zona Rural...")
    st.info("💡 Preencha estes dados primeiro.")

# --- TÍTULO PRINCIPAL ---
st.title("🚓 Gerador de Relatório Policial")
st.markdown("##### Ferramenta Oficial de Padronização - PCPE")

# --- ABAS DE NAVEGAÇÃO ---
tab1, tab2, tab3, tab4 = st.tabs(["👤 Dados do Alvo", "📝 Relato da Diligência", "📸 Evidências", "👮 Equipe Responsável"])

with tab1:
    col1, col2 = st.columns(2)
    alvo_nome = col1.text_input("Nome Completo do Alvo:", "ALEX DO CARMO CORREIA")
    nascimento = col2.text_input("Data de Nascimento:", "15/04/2004")
    
    col3, col4 = st.columns(2)
    alvo_docs = col3.text_input("Documentação (CPF/RG):", "CPF: ... | RG: ...")
    advogado = col4.text_input("Advogado Presente:", "Dr. Adevaldo...")
    
    testemunha = st.text_input("Testemunha / Acompanhante:", "Sra. Marilene...")

with tab2:
    st.markdown("### Descrição dos Fatos")
    st.caption("O texto será formatado automaticamente com recuo e espaçamento 1.5.")
    texto_relato = st.text_area("Digite o relato detalhado:", height=350, 
        placeholder="Em cumprimento à ordem judicial...")

with tab3:
    st.markdown("### Anexo Fotográfico")
    fotos = st.file_uploader("Arraste as fotos para cá", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])
    if fotos:
        st.success(f"{len(fotos)} fotos anexadas.")

with tab4:
    st.markdown("### Responsáveis pela Diligência")
    st.caption("Adicione quantos agentes forem necessários.")
    
    # Loop para criar campos de agentes dinamicamente
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        nome = c1.text_input(f"Nome do Agente {i+1}", key=f"nome_{i}")
        cargo = c2.text_input(f"Cargo/Matrícula {i+1}", key=f"cargo_{i}", value="Investigador de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
        st.markdown("---")
    
    b1, b2 = st.columns(2)
    b1.button("➕ Adicionar Agente", on_click=add_agente)
    b2.button("➖ Remover Agente", on_click=remove_agente)

# --- BOTÃO DE GERAÇÃO ---
st.markdown("---")
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    gerar = st.button("🚀 GERAR RELATÓRIO FINAL (.DOCX)", type="primary")

# --- LÓGICA DE GERAÇÃO DO ARQUIVO ---
if gerar:
    doc = Document()
    
    # 1. Configurar Margens
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)

    # 2. Cabeçalho Simples
    p = doc.add_paragraph()
    r = p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    # 3. Título
    p = doc.add_paragraph()
    r = p.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 4. Dados Técnicos
    def add_dado(label, valor):
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(str(valor))
        aplicar_estilo(p, 11, espaco_depois=2)

    data_formatada = data_doc.strftime("%d de %B de %Y")
    add_dado("OPJ", f"\"{opj}\"")
    add_dado("PROCESSO nº", processo)
    add_dado("DATA", data_formatada)
    add_dado("HORA", str(hora_doc))
    add_dado("LOCAL", local)
    doc.add_paragraph()

    # 5. Seção Alvo
    p = doc.add_paragraph()
    aplicar_estilo(p.add_run("DO ALVO E TESTEMUNHAS"), negrito=True, espaco_depois=6)
    
    add_dado("ALVO", f"{alvo_nome} | {alvo_docs}")
    add_dado("Nascimento", nascimento)
    add_dado("ADVOGADO", advogado)
    add_dado("TESTEMUNHA", testemunha)
    doc.add_paragraph()

    # 6. Relato (Formatado)
    p = doc.add_paragraph()
    aplicar_estilo(p.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO"), negrito=True, espaco_depois=6)
    
    paragrafos = texto_relato.split('\n')
    for par in paragrafos:
        if par.strip():
            p_novo = doc.add_paragraph(par)
            aplicar_estilo(p_novo, 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)

    # 7. Fotos
    if fotos:
        for f in fotos:
            doc.add_page_break()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(f, width=Inches(5.5))
            p_leg = doc.add_paragraph()
            p_leg.add_run(f"Registro Fotográfico: {f.name}")
            aplicar_estilo(p_leg, 9, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 8. Assinaturas Dinâmicas
    doc.add_paragraph()
    doc.add_paragraph()
    
    for agente in agentes_dados:
        if agente['nome']: # Só imprime se tiver nome
            doc.add_paragraph() # Espaço entre assinaturas
            p_sig = doc.add_paragraph()
            p_sig.add_run(f"__________________________________________\n{agente['nome']}\n{agente['cargo']}")
            aplicar_estilo(p_sig, 11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    # Download
    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("📥 BAIXAR SEU RELATÓRIO PRONTO", bio.getvalue(), "Relatorio_PCPE_Pro.docx", type="primary")
