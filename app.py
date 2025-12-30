import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador Universal PCPE", layout="wide", page_icon="🚓")

# --- ESTILO CSS ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: bold;}
    .tag-foto {
        background-color: #e3f2fd; border: 1px solid #1565c0; color: #1565c0;
        padding: 2px 8px; border-radius: 4px; font-weight: bold; font-family: monospace;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE FORMATAÇÃO (ARIAL / ABNT) ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    # Configura Fonte
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

    # Configura Parágrafo
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois)
    p_format.line_spacing = entrelinhas
    if recuo > 0: p_format.first_line_indent = Cm(recuo)
    if alinhamento is not None: paragrafo.alignment = alinhamento

# --- GERENCIAMENTO DE AGENTES ---
if 'num_agentes' not in st.session_state:
    st.session_state.num_agentes = 1

def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# --- BARRA LATERAL (CONFIGURAÇÕES GERAIS) ---
with st.sidebar:
    st.header("1. Cabeçalho do Documento")
    # Título editável para servir para qualquer crime
    titulo_doc = st.text_input("Título do Relatório:", value="RELATÓRIO DE INVESTIGAÇÃO")
    
    st.markdown("---")
    opj = st.text_input("OPJ (Opcional):", placeholder="Ex: INTERCEPTUM")
    natureza = st.text_input("Natureza do Fato:", placeholder="Ex: Homicídio Doloso, Tráfico...")
    processo = st.text_input("Nº Processo / BO / IP:", placeholder="0000000-00.2025...")
    
    c1, c2 = st.columns(2)
    data_doc = c1.date_input("Data:")
    hora_doc = c2.time_input("Hora:")
    
    local = st.text_input("Local do Fato/Diligência:", placeholder="Rua, Bairro, Cidade...")

# --- ÁREA PRINCIPAL ---
st.title("🚓 Gerador de Relatório Policial (Genérico)")
st.caption("Preencha apenas o que for necessário. Campos em branco não aparecerão no documento.")

# --- ABAS DE PREENCHIMENTO ---
tab_env, tab_texto, tab_fotos, tab_equipe = st.tabs(["👥 Envolvidos", "📝 Texto do Relatório", "📸 Fotos", "👮 Equipe"])

with tab_env:
    st.subheader("Dados das Partes")
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("##### 🔴 Alvo / Suspeito (Se houver)")
        alvo_nome = st.text_input("Nome do Alvo:")
        alvo_alcunha = st.text_input("Vulgo / Alcunha:")
        alvo_docs = st.text_input("CPF / RG (Alvo):")
        alvo_nasc = st.text_input("Nascimento / Idade:")
    
    with col_b:
        st.markdown("##### 🔵 Vítima / Testemunha")
        vitima_nome = st.text_input("Nome da Vítima:")
        testemunha_nome = st.text_input("Nome da Testemunha:")
        advogado_nome = st.text_input("Advogado Presente:")

with tab_texto:
    st.subheader("Redação")
    st.info("💡 Dica: Use **[FOTO1]**, **[FOTO2]** no meio do texto para inserir as imagens automaticamente nessa posição.")
    texto_relato = st.text_area("Descreva a diligência ou investigação:", height=400, placeholder="No dia tal, a equipe deslocou-se...")

# Variável para guardar as fotos
fotos_carregadas = []

with tab_fotos:
    st.subheader("Upload de Evidências")
    fotos_carregadas = st.file_uploader("Selecione as imagens", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])
    
    if fotos_carregadas:
        st.markdown("---")
        st.write("##### Códigos para inserção no texto:")
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i % 5]:
                st.image(f, width=80)
                st.markdown(f"<span class='tag-foto'>[FOTO{i+1}]</span>", unsafe_allow_html=True)

with tab_equipe:
    st.subheader("Responsáveis")
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        # Agora começa vazio para você preencher
        nome = c1.text_input(f"Nome do Agente {i+1}", key=f"n{i}", placeholder="Nome Completo")
        cargo = c2.text_input(f"Cargo/Matrícula {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
    
    st.button("➕ Adicionar Policial", on_click=add_agente)
    st.button("➖ Remover Policial", on_click=remove_agente)

# --- GERAÇÃO DO DOCUMENTO ---
st.markdown("---")
if st.button("🚀 GERAR RELATÓRIO OFICIAL", type="primary"):
    doc = Document()
    
    # 1. Margens ABNT/Oficial
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7); sec.right_margin = Inches(0.7)

    # 2. Cabeçalho (Padrão PCPE Simples)
    p = doc.add_paragraph()
    p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    # 3. Título (O que você digitou na barra lateral)
    p = doc.add_paragraph()
    p.add_run(titulo_doc.upper())
    aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 4. Bloco de Dados Iniciais (Só adiciona se tiver texto)
    def add_linha(rotulo, valor):
        if valor: # Só imprime se o usuário digitou algo
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
    
    doc.add_paragraph() # Espaço

    # 5. Seção Envolvidos (Genérica)
    # Verifica se tem algum dado preenchido para criar o título
    tem_dados = any([alvo_nome, vitima_nome, testemunha_nome, advogado_nome])
    
    if tem_dados:
        p = doc.add_paragraph()
        p.add_run("DOS ENVOLVIDOS")
        aplicar_estilo(p, negrito=True, espaco_depois=6)

        if alvo_nome:
            texto_alvo = alvo_nome
            if alvo_alcunha: texto_alvo += f" (Vulgo: {alvo_alcunha})"
            if alvo_docs: texto_alvo += f" | {alvo_docs}"
            add_linha("ALVO/SUSPEITO", texto_alvo)
            add_linha("DADOS DO ALVO", alvo_nasc)
        
        add_linha("VÍTIMA", vitima_nome)
        add_linha("TESTEMUNHA", testemunha_nome)
        add_linha("ADVOGADO", advogado_nome)
        
        doc.add_paragraph()

    # 6. Texto do Relatório (Com sistema de fotos)
    p = doc.add_paragraph()
    p.add_run("DO RELATO / DILIGÊNCIA")
    aplicar_estilo(p, negrito=True, espaco_depois=6)

    # Lógica de inserção de imagem
    partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
    
    for parte in partes:
        if parte.isdigit():
            idx = int(parte) - 1
            if 0 <= idx < len(fotos_carregadas):
                foto = fotos_carregadas[idx]
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img = p_img.add_run()
                run_img.add_picture(foto, width=Inches(5.5))
                
                p_leg = doc.add_paragraph()
                p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_leg.add_run(f"Figura {idx+1}") # Legenda simples
                aplicar_estilo(p_leg, 9, espaco_depois=12)
        else:
            # Texto normal
            linhas = parte.split('\n')
            for linha in linhas:
                if linha.strip():
                    p_txt = doc.add_paragraph()
                    p_txt.add_run(linha)
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
    st.download_button("📥 BAIXAR RELATÓRIO (.DOCX)", bio.getvalue(), "Relatorio_Generico.docx", type="primary")
