import streamlit as st
from docx import Document
from docx.shared import Inches
import io
from datetime import datetime

# Configurações do site
st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("🔍 Gerador de Relatório Automático")
st.markdown("---")

# --- ENTRADA DE DADOS ---
nome = st.text_input("👤 Nome do Investigador:")
titulo = st.text_input("📋 Título da Investigação:")
relato = st.text_area("📝 Relato dos Fatos:", height=250)

# Upload de múltiplas fotos
fotos = st.file_uploader("📸 Suba as fotos aqui", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])

# --- BOTÃO PARA GERAR ---
if st.button("🚀 GERAR RELATÓRIO AGORA"):
    if not relato or not titulo:
        st.error("❌ Por favor, preencha o Título e o Relato.")
    else:
        # Criando o Word
        doc = Document()
        doc.add_heading('RELATÓRIO DE INVESTIGAÇÃO', 0)
        
        # Cabeçalho organizado
        doc.add_paragraph(f"Investigador: {nome}")
        doc.add_paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}")
        doc.add_heading(f"Caso: {titulo}", level=1)
        
        # Texto do relato
        doc.add_heading("Descrição da Ocorrência", level=2)
        doc.add_paragraph(relato)
        
        # Inserindo fotos
        if fotos:
            doc.add_heading("Evidências Fotográficas", level=2)
            for i, foto in enumerate(fotos):
                doc.add_paragraph(f"Evidência {i+1}:")
                doc.add_picture(foto, width=Inches(5))
                doc.add_paragraph("-" * 30)

        # Preparar para download
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        st.success("✅ Relatório pronto para baixar!")
        st.download_button(
            label="⬇️ BAIXAR RELATÓRIO (.DOCX)",
            data=buffer,
            file_name=f"Relatorio_{titulo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
