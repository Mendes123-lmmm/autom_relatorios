import streamlit as st
import matplotlib.pyplot as plt
import openpyxl
import xlwings as xw
from docx import Document
from docx.shared import Inches
import re
import os
from datetime import datetime
import time  # Para garantir tempo suficiente para salvar o gráfico

def replace_text_keep_format(paragraph, old_text, new_text):
    """Substitui texto em um parágrafo mantendo a formatação."""
    for run in paragraph.runs:
        if old_text in run.text:
            updated_text = re.sub(rf'\b{re.escape(old_text)}\b', new_text, run.text)
            run.text = updated_text

def format_value(value):
    """Formata valores numéricos e datas para o padrão desejado."""
    if isinstance(value, (int, float)):
        return '{:,.2f}'.format(value).replace(',', 'X').replace('.', ',').replace('X', '.')
    elif isinstance(value, datetime):
        return value.strftime('%d/%m/%Y')
    return str(value)

st.title("Gerador de Documentos Automatizado")

uploaded_excel = st.file_uploader("Carregar Arquivo Excel", type=["xlsx"])
uploaded_word = st.file_uploader("Carregar Arquivo Word", type=["docx"])

if uploaded_excel and uploaded_word:
    if st.button("Gerar Documento"):
        # Criar arquivos temporários
        excel_path = "temp_excel.xlsx"
        word_path = "temp_word.docx"
        img_path = "grafico_gerado.png"
        output_word_path = "Documento_Atualizado.docx"

        # Inicializar a barra de progresso
        progress_bar = st.progress(0)

        try:
            with open(excel_path, "wb") as f:
                f.write(uploaded_excel.getbuffer())
            progress_bar.progress(10)

            with open(word_path, "wb") as f:
                f.write(uploaded_word.getbuffer())
            progress_bar.progress(20)

            # Carregar o Excel
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheet_cliente = wb['Cliente']
            sheet_max = wb['Específicos']
            sheet_fonte = wb['Fonte']
            sheet_graf = wb["ChartData"]

            # Ler os dados das células
            data = {
                "TAG1": sheet_max['A3'].value,
                "TAG2": sheet_fonte['C52'].value,
                "NOME1": sheet_cliente['B1'].value,
                "NOME2": sheet_cliente['B2'].value,
                "NOME3": sheet_cliente['B3'].value,
            }

            # Formatar os valores
            for key in data:
                data[key] = format_value(data[key])
            progress_bar.progress(40)

            # Coletar os dados do gráfico
            x_values = []
            y_values = []

            for row in sheet_graf.iter_rows(min_row=1, max_row=sheet_graf.max_row, min_col=1, max_col=2, values_only=True):
                if row[0] is not None and row[1] is not None:
                    x_values.append(row[0])
                    y_values.append(row[1])

            # Criar e salvar o gráfico
            plt.figure(figsize=(6, 4))
            plt.plot(x_values, y_values, marker="o", linestyle="-", color="b")
            plt.xlabel("Tempo (ms)")
            plt.ylabel("Intensidade (%)")
            plt.grid(True)
            plt.savefig(img_path)
            plt.close()
            progress_bar.progress(60)

            # Abrir o documento Word e substituir os textos
            doc = Document(word_path)

            for table in doc.tables:
                for cell in table._cells:
                    for paragraph in cell.paragraphs:
                        for key, value in data.items():
                            replace_text_keep_format(paragraph, key, value)

            for paragraph in doc.paragraphs:
                for key, value in data.items():
                    replace_text_keep_format(paragraph, key, value)

                if 'INSERIR_GRAFICO' in paragraph.text:
                    paragraph.text = paragraph.text.replace('INSERIR_GRAFICO', '')
                    run = paragraph.add_run()
                    run.add_picture(img_path, width=Inches(2.5))
                    break
            progress_bar.progress(80)

            # Salvar o novo documento
            doc.save(output_word_path)

            # Disponibilizar para download
            with open(output_word_path, "rb") as f:
                st.download_button("Baixar Documento Atualizado", f, file_name=f"{data['NOME1']}.docx")
            progress_bar.progress(100)

        finally:
            # Fechar processos e limpar arquivos temporários
            try:
                xw.App().quit()
            except NameError:
                pass

            time.sleep(1)

            for file in [excel_path, word_path, output_word_path, img_path]:
                if os.path.exists(file):
                    try:
                        os.remove(file)
                    except PermissionError:
                        print(f"Arquivo em uso, não pode ser excluído: {file}")
