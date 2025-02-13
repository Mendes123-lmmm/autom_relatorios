import streamlit as st
import os

# Título da aplicação
st.title("Gerador de Documentos Automatizado")

# Caixa de seleção para escolher o tipo de documento
opcao = st.selectbox(
    "Selecione o tipo de documento:",
    ("Selecione uma opção", "Constância de Arco", "Relatório 2", "Relatório 3")
)

# Verifica se uma opção válida foi selecionada
if opcao != "Selecione uma opção":
    # Upload do arquivo Excel
    uploaded_excel = st.file_uploader("Carregue seu arquivo Excel", type=["xlsx"])

    # Upload do arquivo Word (modelo)
    uploaded_word = st.file_uploader("Carregue seu modelo Word", type=["docx"])

    # Botão "Gerar Relatório"
    if uploaded_excel and uploaded_word:
        if st.button("Gerar Relatório"):
            if opcao == "Constância de Arco":
                from Const_Arco import processar_constancia_arco
                processar_constancia_arco(uploaded_excel, uploaded_word)
            elif opcao == "Relatório 2":
                st.write("Funcionalidade para Relatório 2 ainda não implementada.")
            elif opcao == "Relatório 3":
                st.write("Funcionalidade para Relatório 3 ainda não implementada.")