import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re
import google.generativeai as genai

# --- 0. CHAVE API ---
CHAVE_API_GOOGLE = "AIzaSyBCdhqPkOVtQtO9x-pQTABb7X258-Si4VQ"

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador PCPE - IA Integrada", layout="wide", page_icon="🚓")

# --- 2. ESTILO ---
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stTextInput>div>div>input {font-weight: 500;}
    .stTextArea textarea {font-size: 15px; line-height: 1.6; font-family: 'Arial';}
    .tag-foto {background-color: #e3f2fd; border: 1px solid #1565c0; color: #1565c0; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-family: monospace;}
    .sucesso-ia {padding: 10px; background-color: #d4edda; color: #155724; border-radius: 5px; margin-bottom: 10px; border: 1px solid #c3e6cb;}
    </style>
""", unsafe_allow_html=True)

# --- 3. FUNÇÕES ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois)
    p_format.line_spacing = entrelinhas
    if recuo > 0: p_format.first_line_indent = Cm(recuo)
    if alinhamento is not None: paragrafo.alignment = alinhamento

def melhorar_texto_com_ia(texto_bruto):
    try:
        genai.configure(api_key=CHAVE_API_GOOGLE)
        # Tenta o modelo Flash (mais rápido e moderno)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Você é um escrivão de polícia experiente. Reescreva o relato abaixo para um Relatório Oficial de Investigação.
        
        DIRETRIZES:
        1. Corrija rigorosamente a gramática e ortografia.
        2. Utilize linguagem formal, técnica e impessoal (Ex: substitua "eu vi" por "a equipe visualizou").
        3. Mantenha as tags de fotos (ex: [FOTO1]) EXATAMENTE onde estão.
        4. Seja claro, conciso e cronológico.
        5. Não invente informações.
        
        RASCUNHO ORIGINAL:
        {texto_bruto}
        """
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        # Se der erro, tenta listar os modelos disponíveis para ajudar no diagnóstico
        erro_msg = str(e)
        if "404" in erro_msg:
            return f"⚠️ ERRO DE VERSÃO: O Streamlit não atualizou a biblioteca. \n\nSOLUÇÃO: Vá no painel do Streamlit, clique nos 3 pontinhos do app e selecione 'Reboot app'."
        return f"Erro na IA: {erro_msg}"

# --- 4. ESTADO ---
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
if 'texto_final' not in st.session_state: st.session_state.texto_final = ""
def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# --- 5. BARRA LATERAL ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/9203/9203764.png", width=60)
    st.header("Configurações")
    
    # Verifica a versão da biblioteca (Diagnóstico)
    versao_lib = genai.__version__
    if versao_lib < "0.7.0":
        st.error(f"⚠️ Biblioteca Antiga Detectada ({versao_lib})")
        st.info("Por favor, faça o REBOOT do app para atualizar.")
    else:
        st.success(f"✅ Sistema Atualizado (v{versao_lib})")

    st.divider()
    st.subheader("📄 Cabeçalho")
    titulo_doc = st.text_input("Título:", value="RELATÓRIO DE INVESTIGAÇÃO")
    opj = st.text_input("OPJ:", placeholder="Ex: INTERCEPTUM")
    natureza = st.text_input("Natureza:", placeholder="Ex: Homicídio...")
    processo = st.text_input("Nº Processo/BO:", placeholder="0000...")
    c1, c2 = st.columns(2)
    data_doc = c1.date_input("Data:")
    hora_doc = c2.time_input("Hora:")
    local = st.text_input("Local:", placeholder="Endereço completo...")

# --- 6. INTERFACE ---
st.title("🚓 Gerador Policial com IA")

tab_env, tab_texto, tab_fotos, tab_equipe = st.tabs(["👥 Envolvidos", "✨ Relato (IA)", "📸 Fotos", "👮 Equipe"])

# ABA 1: ENVOLVIDOS
with tab_env:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 🔴 Suspeito / Alvo")
        alvo_nome = st.text_input("Nome Alvo:")
        alvo_alcunha = st.text_input("Vulgo/Alcunha:")
        alvo_docs = st.text_input("CPF/RG:")
        alvo_nasc = st.text_input("Nascimento:")
    with c2:
        st.markdown("#### 🔵 Outros")
        vitima_nome = st.text_input("Nome Vítima:")
        testemunha_nome = st.text_input("Testemunha:")
        advogado_nome = st.text_input("Advogado:")

# Variável fotos
fotos_carregadas = []

# ABA 3: FOTOS
with tab_fotos:
    st.subheader("Upload de Imagens")
    fotos_carregadas = st.file_uploader("Selecione fotos", accept_multiple_files=True)
    if fotos_carregadas:
        st.info("Use estes códigos no texto:")
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i%5]:
                st.image(f, width=70)
                st.markdown(f"<span class='tag-foto'>[FOTO{i+1}]</span>", unsafe_allow_html=True)

# ABA 2: TEXTO IA
with tab_texto:
    c_in, c_out = st.columns(2)
    with c_in:
        st.markdown("#### Rascunho")
        rascunho = st.text_area("Digite o relato bruto:", height=400, 
            placeholder="Ex: Chegamos e ele correu [FOTO1]...")
        
        if st.button("✨ MELHORAR TEXTO"):
            if not rascunho:
                st.warning("Escreva algo primeiro!")
            else:
                with st.spinner("A IA está reescrevendo..."):
                    res = melhorar_texto_com_ia(rascunho)
                    st.session_state.texto_final = res
                    st.rerun()

    with c_out:
        st.markdown("#### Texto Final")
        if st.session_state.texto_final:
            if "ERRO" in st.session_state.texto_final:
                st.error(st.session_state.texto_final)
            else:
                st.markdown("<div class='sucesso-ia'>✅ Texto Pronto!</div>", unsafe_allow_html=True)
        texto_oficial = st.text_area("Resultado:", height=400, value=st.session_state.texto_final)

# ABA 4: EQUIPE
with tab_equipe:
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3,2])
        n = c1.text_input(f"Nome {i+1}", key=f"n{i}")
        c = c2.text_input(f"Cargo {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes_dados.append({'nome':n, 'cargo':c})
    c1, c2 = st.columns([1,5])
    c1.button("➕", on_click=add_agente)
    c2.button("➖", on_click=remove_agente)

# --- 7. GERAR DOC ---
st.markdown("---")
if st.button("🚀 BAIXAR RELATÓRIO FINAL", type="primary"):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7); sec.right_margin = Inches(0.7)

    # Cabeçalho
    p = doc.add_paragraph()
    p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()

    # Título
    p = doc.add_paragraph()
    p.add_run(titulo_doc.upper())
    aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # Dados
    def add_line(k, v):
        if v:
            p = doc.add_paragraph()
            p.add_run(f"{k}: ").bold = True
            p.add_run(str(v))
            aplicar_estilo(p, 11, espaco_depois=2)
    
    dt = data_doc.strftime("%d/%m/%Y")
    add_line("NATUREZA", natureza)
    add_line("OPJ", opj)
    add_line("REFERÊNCIA", processo)
    add_line("DATA/HORA", f"{dt} às {hora_doc}")
    add_line("LOCAL", local)
    doc.add_paragraph()

    # Envolvidos
    if any([alvo_nome, vitima_nome, testemunha_nome, advogado_nome]):
        p = doc.add_paragraph()
        p.add_run("DOS ENVOLVIDOS")
        aplicar_estilo(p, negrito=True, espaco_depois=6)
        if alvo_nome:
            txt = alvo_nome
            if alvo_alcunha: txt += f" (Vulgo: {alvo_alcunha})"
            if alvo_docs: txt += f" | {alvo_docs}"
            add_line("ALVO", txt)
            if alvo_nasc: add_line("NASCIMENTO", alvo_nasc)
        add_line("VÍTIMA", vitima_nome)
        add_line("TESTEMUNHA", testemunha_nome)
        add_line("ADVOGADO", advogado_nome)
        doc.add_paragraph()

    # Relato
    p = doc.add_paragraph()
    p.add_run("DO RELATO")
    aplicar_estilo(p, negrito=True, espaco_depois=6)
    
    txt_uso = texto_oficial if texto_oficial else rascunho
    parts = re.split(r'\[FOTO(\d+)\]', txt_uso)
    
    for part in parts:
        if part.isdigit():
            idx = int(part) - 1
            if 0 <= idx < len(fotos_carregadas):
                f = fotos_carregadas[idx]
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(f, width=Inches(5.5))
                p2 = doc.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.add_run(f"Figura {idx+1}")
                aplicar_estilo(p2, 9, espaco_depois=12)
        else:
            lines = part.split('\n')
            for line in lines:
                if line.strip():
                    p = doc.add_paragraph()
                    p.add_run(line)
                    aplicar_estilo(p, 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)

    # Assinaturas
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
    st.download_button("📥 BAIXAR DOCX", bio.getvalue(), "Relatorio_Final.docx", type="primary")
