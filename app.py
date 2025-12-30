import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import io
import re
import google.generativeai as genai

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Relatórios Policiais com IA", layout="wide", page_icon="🚓")

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
    .sucesso-ia {
        padding: 10px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        margin-bottom: 10px;
        border: 1px solid #c3e6cb;
    }
    </style>
""", unsafe_allow_html=True)

# --- 3. FUNÇÕES UTILITÁRIAS ---

def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo=0):
    """Aplica formatação ABNT/Policial no parágrafo do Word."""
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
    
    if recuo > 0:
        p_format.first_line_indent = Cm(recuo)
    
    if alinhamento is not None:
        paragrafo.alignment = alinhamento

def melhorar_texto_com_ia(texto_bruto, api_key):
    """Usa o Google Gemini para reescrever o texto policialmente."""
    if not api_key:
        return "⚠️ ERRO: Configure a API Key na barra lateral primeiro."
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Você é um escrivão de polícia experiente. Reescreva o relato abaixo para um Relatório Oficial de Investigação.
        
        DIRETRIZES:
        1. Corrija rigorosamente a gramática e ortografia.
        2. Utilize linguagem formal, técnica e impessoal (Ex: substitua "eu vi" por "a equipe visualizou").
        3. Mantenha as tags de fotos (ex: [FOTO1], [FOTO2]) EXATAMENTE onde estão.
        4. Seja claro, conciso e cronológico.
        5. Não invente informações que não estejam no rascunho.
        
        RASCUNHO ORIGINAL:
        {texto_bruto}
        """
        
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Erro na IA: {str(e)}"

# --- 4. GERENCIAMENTO DE ESTADO (SESSION STATE) ---
if 'num_agentes' not in st.session_state: st.session_state.num_agentes = 1
if 'texto_final' not in st.session_state: st.session_state.texto_final = ""

def add_agente(): st.session_state.num_agentes += 1
def remove_agente(): 
    if st.session_state.num_agentes > 1: st.session_state.num_agentes -= 1

# --- 5. BARRA LATERAL (CONFIGURAÇÕES) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/9203/9203764.png", width=60)
    st.header("🔐 Configuração")
    
    # Tenta pegar chave dos Segredos ou pede input
    if 'GOOGLE_API_KEY' in st.secrets:
        api_key = st.secrets['GOOGLE_API_KEY']
        st.success("🔑 API Key ativa (Segredos)")
    else:
        # AQUI ESTAVA O ERRO: Corrigido de help(...) para help="..."
        api_key = st.text_input("Cole sua Google API Key:", type="password", help="Pegue em aistudio.google.com")

    st.divider()
    st.subheader("📄 Cabeçalho do Doc")
    titulo_doc = st.text_input("Título:", value="RELATÓRIO DE INVESTIGAÇÃO")
    opj = st.text_input("OPJ (Opj):", placeholder="Ex: INTERCEPTUM")
    natureza = st.text_input("Natureza:", placeholder="Ex: Homicídio, Tráfico...")
    processo = st.text_input("Nº Processo/BO:", placeholder="0000000-00.2025...")
    
    c1, c2 = st.columns(2)
    data_doc = c1.date_input("Data:")
    hora_doc = c2.time_input("Hora:")
    local = st.text_input("Local:", placeholder="Endereço completo...")

# --- 6. INTERFACE PRINCIPAL ---
st.title("🚓 Gerador de Relatório Policial Integrado")
st.caption("Preencha os dados, use a IA para melhorar o texto e baixe o DOCX formatado.")

# Abas de Navegação
tab_env, tab_texto, tab_fotos, tab_equipe = st.tabs(["👥 Envolvidos", "✨ Relato com IA", "📸 Evidências", "👮 Equipe"])

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

# Variável global de fotos
fotos_carregadas = []

# ABA 3: FOTOS (Precisa vir antes do texto para mostrar os códigos)
with tab_fotos:
    st.subheader("Upload de Imagens")
    fotos_carregadas = st.file_uploader("Selecione as fotos da diligência", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])
    
    if fotos_carregadas:
        st.info("💡 Use os códigos abaixo no seu texto para posicionar as fotos.")
        cols = st.columns(5)
        for i, f in enumerate(fotos_carregadas):
            with cols[i % 5]:
                st.image(f, width=80)
                st.markdown(f"<span class='tag-foto'>[FOTO{i+1}]</span>", unsafe_allow_html=True)
                st.caption(f"{f.name[:10]}...")

# ABA 2: TEXTO COM IA
with tab_texto:
    col_rascunho, col_final = st.columns(2)
    
    with col_rascunho:
        st.markdown("#### 1. Rascunho (Entrada)")
        st.caption("Escreva os fatos de forma simples. Indique fotos com [FOTO1].")
        rascunho = st.text_area("Digite o relato bruto:", height=450, 
            placeholder="Ex: Chegamos no local às 10h. O indivíduo correu para os fundos. [FOTO1] Encontramos a droga no quarto. [FOTO2]...")
        
        if st.button("✨ REESCREVER COM IA"):
            if not api_key:
                st.error("⚠️ Configure a API Key na barra lateral!")
            elif not rascunho:
                st.warning("⚠️ Escreva algo no rascunho primeiro.")
            else:
                with st.spinner("A IA está trabalhando no seu texto..."):
                    texto_melhorado = melhorar_texto_com_ia(rascunho, api_key)
                    st.session_state.texto_final = texto_melhorado
                    st.rerun() # Atualiza a tela

    with col_final:
        st.markdown("#### 2. Texto Oficial (Final)")
        st.caption("Texto formatado pela IA. Este será usado no documento.")
        
        if st.session_state.texto_final:
            st.markdown(f"<div class='sucesso-ia'>✅ Texto reescrito com sucesso!</div>", unsafe_allow_html=True)
            
        texto_oficial = st.text_area("Edite se necessário:", height=450, value=st.session_state.texto_final)

# ABA 4: EQUIPE
with tab_equipe:
    st.subheader("Responsáveis pela Diligência")
    agentes_dados = []
    for i in range(st.session_state.num_agentes):
        c1, c2 = st.columns([3, 2])
        nome = c1.text_input(f"Nome do Agente {i+1}", key=f"n{i}")
        cargo = c2.text_input(f"Cargo/Matrícula {i+1}", key=f"c{i}", value="Agente de Polícia")
        agentes_dados.append({'nome': nome, 'cargo': cargo})
    
    b1, b2 = st.columns([1, 5])
    b1.button("➕ Adicionar", on_click=add_agente)
    b2.button("➖ Remover", on_click=remove_agente)

# --- 7. GERAÇÃO DO DOCUMENTO ---
st.markdown("---")
col_gerar, _ = st.columns([1, 2])

with col_gerar:
    if st.button("🚀 GERAR RELATÓRIO FINAL (.DOCX)", type="primary"):
        # Inicializa Documento
        doc = Document()
        
        # 1. Margens (Padrão ABNT/Oficial)
        sec = doc.sections[0]
        sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.7); sec.right_margin = Inches(0.7)

        # 2. Cabeçalho Simples (Sem Logo)
        p = doc.add_paragraph()
        p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
        aplicar_estilo(p, 10, True, WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()

        # 3. Título
        p = doc.add_paragraph()
        p.add_run(titulo_doc.upper())
        aplicar_estilo(p, 12, True, WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

        # 4. Metadados
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

        # 5. Envolvidos (Só mostra se tiver dados)
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

        # 6. Relato (Processamento de Texto + Fotos)
        p = doc.add_paragraph()
        p.add_run("DO RELATO")
        aplicar_estilo(p, negrito=True, espaco_depois=6)

        # Usa o texto oficial se existir, senão usa o rascunho
        texto_para_usar = texto_oficial if texto_oficial else rascunho
        
        # Divide o texto pelas tags de foto [FOTOX]
        partes = re.split(r'\[FOTO(\d+)\]', texto_para_usar)
        
        for parte in partes:
            if parte.isdigit():
                # É um número de foto
                idx = int(parte) - 1
                if 0 <= idx < len(fotos_carregadas):
                    f = fotos_carregadas[idx]
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(f, width=Inches(5.5))
                    
                    p_leg = doc.add_paragraph()
                    p_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_leg.add_run(f"Figura {idx+1}")
                    aplicar_estilo(p_leg, 9, espaco_depois=12)
            else:
                # É texto
                linhas = parte.split('\n')
                for linha in linhas:
                    if linha.strip():
                        p_txt = doc.add_paragraph()
                        p_txt.add_run(linha)
                        # Aplica formatação policial: Justificado, 1.5 linhas, Recuo 1.25cm
                        aplicar_estilo(p_txt, 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, entrelinhas=1.5, espaco_depois=6, recuo=1.25)

        # 7. Assinaturas
        doc.add_paragraph(); doc.add_paragraph()
        for ag in agentes_dados:
            if ag['nome']:
                doc.add_paragraph()
                p = doc.add_paragraph()
                p.add_run(f"___________________________\n{ag['nome']}\n{ag['cargo']}")
                aplicar_estilo(p, 11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

        # 8. Salvar e Baixar
        bio = io.BytesIO()
        doc.save(bio)
        st.balloons()
        st.download_button("📥 DOWNLOAD .DOCX", bio.getvalue(), "Relatorio_Policial_IA.docx", type="primary")
