import os  # Adicione esta linha no início do arquivo
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from fpdf import FPDF
from PIL import Image
import tempfile

# Função para gerar o relatório em PDF
def gerar_relatorio_pdf(df, atividades_sem_predecessora, atividades_atrasadas, caminho_critico, curva_s_img):
    pdf = FPDF()

    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Relatório do Projeto", ln=True, align="C")

    # Salvar temporariamente a imagem da Curva S
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as temp_image:
        img_filename = temp_image.name
        curva_s_img.seek(0)
        with open(img_filename, 'wb') as f:
            f.write(curva_s_img.read())

    # Adicionar a imagem da Curva S no PDF
    pdf.cell(200, 10, txt="Curva S", ln=True)
    pdf.image(img_filename, x=10, y=30, w=190)

    pdf.ln(180)
    pdf.cell(200, 10, txt="Caminho Crítico", ln=True)
    for atividade in caminho_critico:
        pdf.cell(200, 10, txt=atividade, ln=True)

    pdf.ln(10)
    pdf.cell(200, 10, txt="Atividades Sem Predecessoras", ln=True)
    for _, row in atividades_sem_predecessora.iterrows():
        pdf.cell(200, 10, txt=row['Nome da tarefa'], ln=True)

    pdf.ln(10)
    pdf.cell(200, 10, txt="Atividades Atrasadas", ln=True)
    for _, row in atividades_atrasadas.iterrows():
        pdf.cell(200, 10, txt=row['Nome da tarefa'], ln=True)

    output = io.BytesIO()
    pdf.output(output, 'S').encode('latin1')  # Corrigido para salvar no buffer de memória
    output.seek(0)
    
    # Verificar se o arquivo temporário ainda existe antes de removê-lo
    if os.path.exists(img_filename):
        os.remove(img_filename)

    return output

# Resto do código continua igual...
