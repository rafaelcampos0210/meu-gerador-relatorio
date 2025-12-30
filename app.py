import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re  # Biblioteca para encontrar as tags [FOTO1] no texto

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Relatório Policial Inteligente", layout="wide", page_icon="🚓")

# --- ESTILO CSS ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1 {color: #1f2c56;}
    .stButton>button {width: 100%; border-radius: 5px; height: 3em; font-weight: bold;}
    .tag-foto {
        background-color: #e0f7fa; 
        border: 1px solid #006064; 
        color: #006064; 
        padding: 2px 6px; 
        border-radius: 4px; 
        font-family: monospace; 
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- FUNÇÃO DE FORMATAÇÃO ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')

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

def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# --- BARRA LATERAL ---
with st.sidebar:
    st.title("Configurações")
    opj = st.text_input("OPJ:", "INTERCEPTUM")
    processo = st.text_input("Processo:", "0002343-02.2025.8.17.3410")
    data_doc = st.date_input("Data:")
    hora_doc = st.time_input("Hora:")
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural...")

# --- TÍTULO ---
st.title("🚓 Gerador com Inserção Inteligente de Fotos")

# --- ABAS ---
tab_alvo, tab_fotos, tab_relato, tab_equipe = st.tabs(["1. Alvo", "2. Upload de Fotos (Importante)", "3. Relato", "4. Equipe"])

# Variável global para armazenar as fotos carregadas
fotos_carregadas = []

with tab_alvo:
    c1, c2 = st.columns(2)
    alvo_nome = c1.text_input("Nome Alvo:", "ALEX DO CARMO CORREIA")
    alvo_docs = c2.text_input("Docs:", "CPF: ...")
    c3, c4 = st.columns(2)
    nascimento = c3.text_input("Nascimento:", "15/04/2004")
    advogado = c4.text_input("Advogado:", "Dr. Adevaldo...")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene...")

with tab_fotos:
    st.info("📸 Faça o upload das fotos aqui. O sistema gerará um CÓDIGO para cada uma.")
    fotos_carregadas = st.file_uploader("Selecione as imagens", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])
    
    if fotos_carregadas:
        st.write("---")
        st.subheader("📋 Códigos para usar no texto:")
        cols = st.columns(4)
        for i, foto in enumerate(fotos_carregadas):
            with cols[i % 4]:
                # Mostra a imagem pequena e o código dela
                st.image(foto, width=100)
                st.markdown(f"Use: <span class='tag-foto'>[FOTO{i+1}]</span>", unsafe_allow_html=True)
                st.caption(f"Nome: {foto.name}")

with tab_relato:
    st.subheader("Redação do Relatório")
    st.markdown("""
    **Como inserir fotos:**
    Escreva seu texto normalmente. Onde quiser uma imagem, digite o código dela (ex: `[FOTO1]`).
    
    *Exemplo:*
    > *A equipe entrou na residência. **[FOTO1]** No quarto, foi encontrada uma arma em cima da cama. **[FOTO2]** O suspeito foi conduzido...*
    """)
    texto_relato = st.text_area("Digite o relato:", height=400, placeholder="Digite o texto aqui...")

with tab_equipe:
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        nome = c1.text_input(f"Nome {i+1}", key=f"n{i}")
        cargo = c2.text_input(f"Cargo {i+1}", key=f"c{i}", value="Investigador de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
    st.button("➕ Add Agente", on_click=add_agente)

# --- BOTÃO GERAR ---
st.markdown("---")
if st.button("🚀 GERAR RELATÓRIO CUSTOMIZADO", type="primary"):
    doc = Document()
    
    # 1. Margens
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7); sec.right_margin = Inches(0.7)

    # 2. Cabeçalho (Sem Logo/Rodapé conforme preferência anterior)
    p = doc.add_paragraph()
    p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    # 3. Título e Dados
    p = doc.add_paragraph()
    p.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    def add_dado(label, valor):
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(str(valor))
        aplicar_estilo(p, 11, espaco_depois=2)

    data_fmt = data_doc.strftime("%d/%m/%Y")
    add_dado("OPJ", opj); add_dado("PROCESSO", processo)
    add_dado("DATA/HORA", f"{data_fmt} às {hora_doc}"); add_dado("LOCAL", local)
    doc.add_paragraph()

    # 4. Alvo
    p = doc.add_paragraph()
    p.add_run("DO ALVO E TESTEMUNHAS")
    aplicar_estilo(p, negrito=True, espaco_depois=6)
    add_dado("ALVO", f"{alvo_nome} | {alvo_docs}")
    add_dado("DADOS", f"Nasc: {nascimento} | Adv: {advogado}")
    add_dado("TESTEMUNHA", testemunha)
    doc.add_paragraph()

    # 5. DILIGÊNCIA (A Lógica Mágica de Inserção)
    p = doc.add_paragraph()
    p.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    aplicar_estilo(p, negrito=True, espaco_depois=6)

    # Divide o texto procurando por padrões [FOTO1], [FOTO2]...
    # O regex r'\[FOTO(\d+)\]' separa o texto e captura o número da foto
    partes = re.split(r'\[FOTO(\d+)\]', texto_relato)
    
    # O 'partes' vai ser algo como: ["Texto antes", "1", "Texto depois", "2", "Final"]
    
    for i, parte in enumerate(partes):
        # Se for um número (que veio do regex), é hora de por a foto
        if parte.isdigit():
            idx = int(parte) - 1 # Converte "1" para índice 0
            if 0 <= idx < len(fotos_carregadas):
                foto_arquivo = fotos_carregadas[idx]
                
                # Inserir Foto
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img = p_img.add_run()
                run_img.add_picture(foto_arquivo, width=Inches(5.5))
                
                # Inserir Legenda
                p_leg = doc.add_paragraph()
                p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_leg.add_run(f"Figura {idx+1}: {foto_arquivo.name}")
                aplicar_estilo(p_leg, 9, espaco_depois=12)
            else:
                # Caso o usuário digite [FOTO99] e não exista
                p_erro = doc.add_paragraph()
                p_erro.add_run(f"[ERRO: Imagem {parte} não encontrada]").font.color.rgb = None
                aplicar_estilo(p_erro, 11, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

        # Se não for número, é texto normal
        else:
            # Processa o texto para manter parágrafos se houver quebras de linha
            sub_paragrafos = parte.split('\n')
            for sub_p in sub_paragrafos:
                if sub_p.strip():
                    p_texto = doc.add_paragraph()
                    p_texto.add_run(sub_p)
                    aplicar_estilo(p_texto, 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)

    # 6. Assinaturas
    doc.add_paragraph(); doc.add_paragraph()
    for ag in agentes_dados:
        if ag['nome']:
            doc.add_paragraph()
            p = doc.add_paragraph()
            p.add_run(f"___________________________\n{ag['nome']}\n{ag['cargo']}")
            aplicar_estilo(p, 11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    bio = io.BytesIO()
    doc.save(bio)
    st.balloons()
    st.download_button("📥 BAIXAR DOCX COMPLETO", bio.getvalue(), "Relatorio_Com_Fotos_No_Texto.docx", type="primary")
