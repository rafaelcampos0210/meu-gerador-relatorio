import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import io

# --- FUNÇÃO DE ESTILO AVANÇADA ---
def aplicar_estilo(paragrafo, tamanho=11, negrito=False, alinhamento=None, espaco_depois=0, entrelinhas=1.0, recuo_primeira_linha=0):
    """
    Função mestre para controlar cada milímetro do texto.
    """
    # 1. Fonte Arial
    paragrafo.style.font.name = 'Arial'
    paragrafo.style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial')
    
    # 2. Configurações de Parágrafo
    p_format = paragrafo.paragraph_format
    p_format.space_after = Pt(espaco_depois) # Espaço em branco DEPOIS do parágrafo
    p_format.line_spacing = entrelinhas      # Distância entre as linhas do mesmo parágrafo
    
    if recuo_primeira_linha > 0:
        p_format.first_line_indent = Cm(recuo_primeira_linha) # Aquele recuo clássico de início de frase

    if alinhamento is not None:
        paragrafo.alignment = alinhamento

    # 3. Aplica estilo a todos os 'runs' (trechos) do parágrafo
    for run in paragrafo.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(tamanho)
        run.bold = negrito

st.set_page_config(page_title="Gerador PCPE - Formatação Exata", layout="centered")
st.title("🚓 Gerador de Relatório (Formatação ABNT/Policial)")
st.markdown("Este modelo aplica espaçamento 1.5 no texto e recuos de parágrafo.")

# --- FORMULÁRIO ---
with st.form("form_formatacao"):
    st.subheader("1. Cabeçalho")
    col1, col2 = st.columns(2)
    with col1:
        opj = st.text_input("OPJ:", "INTERCEPTUM")
        processo = st.text_input("Processo:", "0002343-02.2025.8.17.3410")
    with col2:
        data = st.text_input("Data:", "22 de dezembro de 2025")
        hora = st.text_input("Hora:", "14h23")
    local = st.text_input("Local:", "Sítio Salvador, nº 360, Zona Rural, Vertente do Lério/PE")

    st.subheader("2. Dados do Alvo")
    alvo_nome = st.text_input("Nome:", "ALEX DO CARMO CORREIA")
    alvo_docs = st.text_input("Docs (CPF/RG):", "CPF: 167.476.854-07 | RG: 8.979.947-9 SDS/PE")
    nascimento = st.text_input("Nascimento:", "15/04/2004")
    advogado = st.text_input("Advogado:", "Dr. Adevaldo do Nascimento Barbosa (OAB/PE 47.508)")
    testemunha = st.text_input("Testemunha:", "Sra. Marilene Lima do Carmo Correia (Genitora)")

    st.subheader("3. Texto da Diligência")
    st.info("O sistema aplicará automaticamente recuo na primeira linha e espaçamento 1.5.")
    texto_input = st.text_area("Digite o relato (use Enter para novos parágrafos):", height=300, 
        value="Em cumprimento à ordem judicial expedida pela Vara Criminal competente, as equipes deslocaram-se ao endereço supracitado para fins de busca domiciliar...\n\nA entrada no domicílio foi autorizada judicialmente...")

    st.subheader("4. Finalização")
    fotos = st.file_uploader("Fotos", accept_multiple_files=True)
    responsavel = st.text_input("Responsável:", "Rafael de Albuquerque Campos")
    cargo = st.text_input("Cargo:", "Investigador de Polícia")
    
    gerar = st.form_submit_button("GERAR DOCX FORMATADO")

if gerar:
    doc = Document()
    
    # MARGENS (Padrão do Modelo)
    sec = doc.sections[0]
    sec.top_margin = Inches(0.5)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)

    # 1. CABEÇALHO (Centralizado, Sem Logo, Espaçamento Simples)
    p = doc.add_paragraph()
    r = p.add_run("POLÍCIA CIVIL DE PERNAMBUCO\nDINTER 1-16ª DESEC\nDelegacia de Polícia da 116ª Circunscrição - Surubim")
    aplicar_estilo(p, tamanho=10, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, entrelinhas=1.0, espaco_depois=0)
    
    doc.add_paragraph() # Espaço em branco manual

    # 2. TÍTULO (Espaçamento Simples)
    p = doc.add_paragraph()
    r = p.add_run("RELATÓRIO DE CUMPRIMENTO DE MANDADO DE BUSCA E APREENSÃO DOMICILIAR")
    aplicar_estilo(p, tamanho=12, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, entrelinhas=1.0, espaco_depois=12)

    # 3. DADOS TÉCNICOS (Bloco Compacto - Espaçamento Simples)
    def add_dado(label, valor):
        p = doc.add_paragraph()
        p.add_run(f"{label}: ").bold = True
        p.add_run(valor)
        # Espaçamento 1.0 (Simples) e 2pt depois para não ficar grudado demais, mas compacto
        aplicar_estilo(p, tamanho=11, entrelinhas=1.0, espaco_depois=2)

    add_dado("OPERAÇÃO DE POLÍCIA JUDICIÁRIA (OPJ)", f"\"{opj}\"")
    add_dado("PROCESSO nº", processo)
    add_dado("DATA", data)
    add_dado("HORA", hora)
    add_dado("LOCAL", local)

    doc.add_paragraph() 

    # 4. SEÇÃO ALVO
    p = doc.add_paragraph()
    p.add_run("DO ALVO E TESTEMUNHAS")
    aplicar_estilo(p, negrito=True, espaco_depois=6) # 6pt de espaço após o título

    add_dado("ALVO", f"{alvo_nome} | {alvo_docs}")
    add_dado("Nascimento", nascimento)
    add_dado("ADVOGADO", advogado)
    add_dado("TESTEMUNHA", testemunha)

    doc.add_paragraph()

    # 5. SEÇÃO DILIGÊNCIA (AQUI ESTÁ A MÁGICA DA FORMATAÇÃO DE TEXTO)
    p = doc.add_paragraph()
    p.add_run("DA DILIGÊNCIA E CUMPRIMENTO DO MANDADO")
    aplicar_estilo(p, negrito=True, espaco_depois=6)

    # Processar o texto: Recuo na primeira linha + Espaçamento 1.5 + Espaço entre parágrafos
    paragrafos_texto = texto_input.split('\n')
    for par in paragrafos_texto:
        if par.strip():
            p_novo = doc.add_paragraph(par)
            aplicar_estilo(
                p_novo, 
                tamanho=11, 
                alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, 
                entrelinhas=1.5,        # 1.5 Linhas (Padrão de Texto Jurídico)
                espaco_depois=6,        # Espaço entre um parágrafo e outro
                recuo_primeira_linha=1.25 # Recuo de 1.25cm no início da linha
            )

    # 6. FOTOS
    if fotos:
        for f in fotos:
            doc.add_page_break()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(f, width=Inches(5.5))
            
            p_leg = doc.add_paragraph()
            p_leg.add_run(f"Registro Fotográfico: {f.name}")
            aplicar_estilo(p_leg, tamanho=9, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, espaco_depois=12)

    # 7. ASSINATURA (Centralizada)
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    p_sig = doc.add_paragraph()
    p_sig.add_run(f"__________________________________________\n{responsavel}\n{cargo}")
    aplicar_estilo(p_sig, tamanho=11, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, entrelinhas=1.0)

    # Salvar
    bio = io.BytesIO()
    doc.save(bio)
    st.success("✅ Documento formatado com espaçamentos corrigidos!")
    st.download_button("⬇️ Baixar DOCX", bio.getvalue(), "Relatorio_Formatacao_Total.docx")
