import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador PCPE - Oficial", layout="wide", page_icon="🚓")

# --- 2. ESTILO VISUAL (CSS) ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: 500;}
    .stTextArea textarea {font-size: 15px; line-height: 1.6; font-family: 'Arial';}
    .tag-foto {
        background-color: #e3f2fd; 
        border: 1px solid #1565c0; 
        color: #1565c0; 
        padding: 2px 8px; 
        border-radius: 4px; 
        font-weight: bold; 
        font-family: monospace;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. FUNÇÃO DE FORMATAÇÃO (ABNT/POLICIAL) ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    # Configura Fonte (Arial)
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

    # Configura Parágrafo
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois)
    p_format.line_spacing = entrelinhas
    
    if recuo > 0:
        p_format.first_line_indent = Cm(recuo)
    
    if alinhamento is not None:
        paragrafo.alignment = alinhamento

# --- 4. GERENCIAMENTO DE ESTADO ---
if 'num_agentes' not in st.session_state:
    st.session_state.num_agentes = 1

def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# --- 5. BARRA LATERAL (CONFIGURAÇÕES) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/9203/9203764.png", width=60)
    st.header("Configurações do Documento")
    st.markdown("---")
    
    titulo_doc = st.text_input("Título do Relatório:", value="RELATÓRIO DE INVESTIGAÇÃO")
    opj = st.text_input("OPJ (Opcional):", placeholder="Ex: INTERCEPTUM")
    natureza = st.text_input("Natureza do Fato:", placeholder="Ex: Homicídio, Tráfico...")
    processo = st.text_input("Nº Processo/BO:", placeholder="0000000-00.2025...")
    
    c1, c2 = st.columns(2)
    data_doc = c1.date_input("Data do Doc:")
    hora_doc = c2.time_input("Hora do Doc:")
    
    local = st.text_input("Local:", placeholder="Endereço completo...")

# --- 6. INTERFACE PRINCIPAL ---
st.title("🚓 Gerador de Relatório Policial")
st.caption("Ferramenta padrão para geração de documentos oficiais.")

# Abas
tab_env, tab_relato, tab_fotos, tab_equipe = st.tabs(["👥 Envolvidos", "📝 Relato", "📸 Evidências", "👮 Equipe"])

# ABA 1: ENVOLVIDOS
with tab_env:
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### 🔴 Suspeito / Alvo")
        alvo_nome = st.text_input("Nome do Alvo:")
        alvo_alcunha = st.text_input("Vulgo / Alcunha:")
        alvo_docs = st.text_input("CPF / RG (Alvo):")
        alvo_nasc = st.text_input("Nascimento / Idade:")
    
    with col_b:
        st.markdown("#### 🔵 Outros Envolvidos")
        vitima_nome = st.text_input("Nome da Vítima:")
        testemunha_nome = st.text_input("Nome da Testemunha:")
        advogado_nome = st.text_input("Advogado Presente:")

# Variável de fotos (Global)
fotos_carregadas = []

# ABA 3: FOTOS (Upload)
with tab_fotos:
    st.subheader("Upload de Imagens")
    fotos_carregadas = st.file_uploader("Selecione as imagens", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])
    
    if fotos_carregadas:
        st.info("Copie os códigos abaixo e cole no texto onde a foto deve aparecer.")
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i % 5]:
                st.image(f, width=80)
                st.markdown(f"<span class='tag-foto'>[FOTO{i+1}]</span>", unsafe_allow_html=True)
                st.caption(f"{f.name[:10]}...")

# ABA 2: RELATO (Texto Manual)
with tab_relato:
    st.subheader("Redação")
    st.markdown("""
    **Instruções:**
    1. Digite seu texto normalmente.
    2. Use **[FOTO1]**, **[FOTO2]** para inserir as imagens entre os parágrafos.
    """)
    texto_relato = st.text_area("Descreva a diligência:", height=450, 
        placeholder="Ex: A equipe chegou ao local às 10h. O indivíduo tentou fugir pelos fundos.\n\n[FOTO1]\n\nApós a captura, foi encontrado o material ilícito...")

# ABA 4: EQUIPE
with tab_equipe:
    st.subheader("Responsáveis")
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        nome = c1.text_input(f"Nome do Agente {i+1}", key=f"n{i}")
        cargo = c2.text_input(f"Cargo/Matrícula {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
    
    b1, b2 = st.columns([1, 5])
    b1.button("➕ Adicionar Agente", on_click=add_agente)
    b2.button("➖ Remover Agente", on_click=remove_agente)

# --- 7. GERAÇÃO DO DOCUMENTO ---
st.markdown("---")
col_gerar, _ = st.columns([1, 2])

with col_gerar:
    if st.button("🚀 GERAR RELATÓRIO OFICIAL (.DOCX)", type="primary"):
        doc = Document()
        
        # 1. Margens (Padrão ABNT)
        sec = doc.sections[0]
        sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.7); sec.right_margin = Inches(0.7)

        # 2. Cabeçalho (Sem Logo, apenas texto centralizado)
        p = doc.add_paragraph()
        p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
        aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()

        # 3. Título (Dinâmico)
        p = doc.add_paragraph()
        p.add_run(titulo_doc.upper())
        aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

        # 4. Dados Iniciais
        def add_linha(rotulo, valor):
            if valor:
                p = doc.add_paragraph()
                p.add_run(f"{rotulo}: ").bold = True
                p.add_run(str(valor))
                aplicar_estilo(p, 11, espaco_depois=2)

        data_fmt = data_doc.strftime("%d/%m/%Y")
        add_linha("NATUREZA", natureza)
        add_linha("OPJ", opj)
        add_linha("REFERÊNCIA", processo)
        add_linha("DATA/HORA", f"{data_fmt} às {hora_doc}")
        add_linha("LOCAL", local)
        doc.add_paragraph()

        # 5. Envolvidos (Condicional)
        if any([alvo_nome, vitima_nome, testemunha_nome, advogado_nome]):
            p = doc.add_paragraph()
            p.add_run("DOS ENVOLVIDOS")
            aplicar_estilo(p, negrito=True, espaco_depois=6)
            
            if alvo_nome:
                txt_alvo = alvo_nome
                if alvo_alcunha: txt_alvo += f" (Vulgo: {alvo_alcunha})"
                if alvo_docs: txt_alvo += f" | {alvo_docs}"
                add_linha("ALVO/SUSPEITO", txt_alvo)
                if alvo_nasc: add_linha("NASCIMENTO", alvo_nasc)
            
            add_linha("VÍTIMA", vitima_nome)
            add_linha("TESTEMUNHA", testemunha_nome)
            add_linha("ADVOGADO", advogado_nome)
            doc.add_paragraph()

        # 6. Relato e Fotos (Lógica de Inserção)
        p = doc.add_paragraph()
        p.add_run("DO RELATO")
        aplicar_estilo(p, negrito=True, espaco_depois=6)

        # Divide o texto onde encontrar [FOTO1], [FOTO2], etc.
        partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
        
        for parte in partes:
            if parte.isdigit():
                # É um número de foto
                idx = int(parte) - 1
                if 0 <= idx < len(fotos_carregadas):
                    f = fotos_carregadas[idx]
                    # Adiciona Foto
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(f, width=Inches(5.5))
                    # Adiciona Legenda
                    p_leg = doc.add_paragraph()
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_leg.add_run(f"Figura {idx+1}")
                    aplicar_estilo(p_leg, 9, espaco_depois=12)
            else:
                # É texto normal
                linhas = parte.split('\n')
                for linha in linhas:
                    if linha.strip():
                        p_txt = doc.add_paragraph()
                        p_txt.add_run(linha)
                        # Formatação Profissional: Justificado, Recuo na 1ª linha, Espaço 1.5
                        aplicar_estilo(p_txt, 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)

        # 7. Assinaturas
        doc.add_paragraph(); doc.add_paragraph()
        for ag in agentes_dados:
            if ag['nome']:
                doc.add_paragraph()
                p = doc.add_paragraph()
                p.add_run(f"___________________________\n{ag['nome']}\n{ag['cargo']}")
                aplicar_estilo(p, 11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

        # Download
        bio = io.BytesIO()
        doc.save(bio)
        st.balloons()
        st.download_button("📥 BAIXAR RELATÓRIO (.DOCX)", bio.getvalue(), "Relatorio_Final.docx", type="primary")
