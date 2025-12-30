import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Relatório Policial Pro", layout="wide", page_icon="🚓")

# --- ESTILO CSS ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1 {color: #1f2c56;}
    .stButton>button {width: 100%; border-radius: 5px; height: 3em; font-weight: bold;}
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE FORMATAÇÃO (CORRIGIDA) ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    # 1. Configurações de Fonte (Aplica em todos os trechos do parágrafo)
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

    # 2. Configurações de Parágrafo (Aplica no bloco inteiro)
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois)
    p_format.line_spacing = entrelinhas
    
    if recuo > 0:
        p_format.first_line_indent = Cm(recuo)
    
    if alinhamento is not None:
        paragrafo.alignment = alinhamento

# --- GERENCIAMENTO DE AGENTES ---
if 'num_agentes' not in st.session_state:
    st.session_state.num_agentes = 2

def add_agente():
    st.session_state.num_agentes += 1

def remove_agente():
    if st.session_state.num_agentes > 1:
        st.session_state.num_agentes -= 1

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("Configurações")
    opj = st.text_input("Nome da OPJ:", "INTERCEPTUM")
    processo = st.text_input("Nº Processo:", "0002343-02.2025.8.17.3410")
    data_doc = st.date_input("Data do Documento:")
    hora_doc = st.time_input("Hora do Documento:")
    local = st.text_input("Local da Diligência:", "Sítio Salvador, nº 360, Zona Rural...")

# --- TÍTULO ---
st.title("🚓 Gerador de Relatório Policial")

# --- ABAS ---
tab1, tab2, tab3, tab4 = st.tabs(["👤 Dados do Alvo", "📝 Relato", "📸 Fotos", "👮 Equipe"])

with tab1:
    col1, col2 = st.columns(2)
    alvo_nome = col1.text_input("Nome do Alvo:", "ALEX DO CARMO CORREIA")
    nascimento = col2.text_input("Nascimento:", "15/04/2004")
    col3, col4 = st.columns(2)
    alvo_docs = col3.text_input("Docs (CPF/RG):", "CPF: ...")
    advogado = col4.text_input("Advogado:", "Dr. Adevaldo...")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene...")

with tab2:
    st.markdown("### Descrição")
    texto_relato = st.text_area("Relato:", height=350, value="Em cumprimento à ordem judicial...")

with tab3:
    st.markdown("### Fotos")
    fotos = st.file_uploader("Upload Fotos", accept_multiple_files=True)

with tab4:
    st.markdown("### Agentes")
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        nome = c1.text_input(f"Nome Agente {i+1}", key=f"nome_{i}")
        cargo = c2.text_input(f"Cargo {i+1}", key=f"cargo_{i}", value="Investigador de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
    st.button("➕ Adicionar Agente", on_click=add_agente)

# --- BOTÃO GERAR ---
st.markdown("---")
if st.button("🚀 GERAR RELATÓRIO (.DOCX)", type="primary"):
    doc = Document()
    
    # 1. Margens
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)

    # 2. Cabeçalho
    p = doc.add_paragraph()
    p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    # 3. Título
    p = doc.add_paragraph()
    p.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
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

    # 5. Seção Alvo (CORREÇÃO AQUI)
    p = doc.add_paragraph()
    p.add_run("DO ALVO E TESTEMUNHAS") # Primeiro adiciona o texto
    aplicar_estilo(p, negrito=True, espaco_depois=6) # Depois aplica o estilo no PARÁGRAFO P
    
    add_dado("ALVO", f"{alvo_nome} | {alvo_docs}")
    add_dado("Nascimento", nascimento)
    add_dado("ADVOGADO", advogado)
    add_dado("TESTEMUNHA", testemunha)
    doc.add_paragraph()

    # 6. Relato (CORREÇÃO AQUI)
    p = doc.add_paragraph()
    p.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    aplicar_estilo(p, negrito=True, espaco_depois=6)
    
    paragrafos = texto_relato.split('\n')
    for par in paragrafos:
        if par.strip():
            p_novo = doc.add_paragraph()
            p_novo.add_run(par)
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

    # 8. Assinaturas
    doc.add_paragraph()
    doc.add_paragraph()
    
    for agente in agentes_dados:
        if agente['nome']:
            doc.add_paragraph()
            p_sig = doc.add_paragraph()
            p_sig.add_run(f"__________________________________________\n{agente['nome']}\n{agente['cargo']}")
            aplicar_estilo(p_sig, 11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("📥 BAIXAR RELATÓRIO", bio.getvalue(), "Relatorio_Pro.docx", type="primary")
