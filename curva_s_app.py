import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import streamlit as st
import io
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from fpdf import FPDF
import datetime

# Função para limpar a abreviação dos dias da semana
def clean_weekday_abbreviation(date_str):
    return date_str.split(' ', 1)[1] if isinstance(date_str, str) else date_str

# Função para ler o arquivo Excel e tratar as colunas de data
def read_excel(file):
    df = pd.read_excel(file)
    
    # Limpar as colunas de datas
    df['Início'] = df['Início'].apply(clean_weekday_abbreviation)
    df['Término'] = df['Término'].apply(clean_weekday_abbreviation)
    
    # Converter para datetime
    df['Início'] = pd.to_datetime(df['Início'], format='%d/%m/%y', errors='coerce')
    df['Término'] = pd.to_datetime(df['Término'], format='%d/%m/%y', errors='coerce')
    
    # Tratar a duração (remover "dias" e converter para float)
    df['Duracao'] = df['Duração'].str.extract('(\d+)').astype(float)
    
    return df

# Função para gerar a Curva S e salvar como PNG
def plot_s_curve(timeline, curva_s, start_date):
    semanas = calcular_numero_semana(timeline, start_date)
    
    fig, ax = plt.subplots()
    ax.plot(semanas, curva_s, marker='o', label="Curva S (0 a 100%)")
    ax.axvline(x=semanas[0], color='green', linestyle='--', label="Início do Cronograma")
    
    ax.set_title('Curva S - Progresso Acumulado (0 a 100%)')
    ax.set_xlabel('Número da Semana')
    ax.set_ylabel('Progresso Acumulado (%)')
    ax.set_ylim(0, 100)
    ax.grid(True)

    ax.set_xticks(semanas)
    plt.xticks(rotation=45)

    plt.legend()
    st.pyplot(fig)

    # Salvar o gráfico como PNG temporário
    curva_s_path = "curva_s.png"
    fig.savefig(curva_s_path)
    plt.close(fig)

    return curva_s_path

# Função para gerar o relatório em PDF
def gerar_relatorio_pdf(df, caminho_critico, atividades_sem_predecessora, atividades_atrasadas, curva_s_path):
    pdf = FPDF()

    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="Relatório Detalhado do Projeto", ln=True, align="C")

    # Adicionar Curva S
    pdf.cell(200, 10, txt="Curva S", ln=True)
    pdf.image(curva_s_path, x=10, y=30, w=190)

    # Adicionar caminho crítico
    pdf.cell(200, 10, txt="Caminho Crítico", ln=True)
    for atividade in caminho_critico:
        pdf.cell(200, 10, txt=atividade, ln=True)

    # Adicionar atividades sem predecessoras
    pdf.cell(200, 10, txt="Atividades Sem Predecessoras", ln=True)
    for _, row in atividades_sem_predecessora.iterrows():
        pdf.cell(200, 10, txt=row['Nome da tarefa'], ln=True)

    # Adicionar atividades atrasadas
    if not atividades_atrasadas.empty:
        pdf.cell(200, 10, txt="Atividades Atrasadas", ln=True)
        for _, row in atividades_atrasadas.iterrows():
            pdf.cell(200, 10, txt=row['Nome da tarefa'], ln=True)

    # Salvar o relatório em PDF
    pdf_output = io.BytesIO()
    pdf.output(pdf_output, 'F')
    pdf_output.seek(0)

    # Remover o arquivo temporário de gráfico
    if os.path.exists(curva_s_path):
        os.remove(curva_s_path)
    
    return pdf_output

# Função para calcular o número da semana a partir de uma data inicial
def calcular_numero_semana(timeline, start_date):
    return [(date - start_date).days // 7 + 1 for date in timeline]

# Interface Streamlit
st.title('Gerador de Curva S e Caminho Crítico com Alerta de Atraso e Relatório PDF')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="DD/MM/AAAA")
end_date = st.text_input("Selecione a data final do cronograma (DD/MM/AAAA)", placeholder="DD/MM/AAAA")

if uploaded_file is not None:
    try:
        # Verifique se o arquivo foi carregado com sucesso
        if uploaded_file is not None:
            # Converter as datas de início e final para o formato correto
            start_date = pd.to_datetime(start_date, format='%d/%m/%Y')
            end_date = pd.to_datetime(end_date, format='%d/%m/%Y')

            # Processar o arquivo Excel
            df = read_excel(uploaded_file)
        
            st.write("Dados do cronograma:")
            st.dataframe(df)
        
            # Continuar com o restante da lógica para caminho crítico, Curva S, etc.
            # ...

        else:
            st.error("Nenhum arquivo foi carregado. Por favor, carregue um arquivo Excel válido.")
    
    except ValueError as e:
        st.error(f"Erro ao processar os dados: {e}")
