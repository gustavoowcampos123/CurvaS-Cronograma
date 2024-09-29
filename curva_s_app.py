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

# Interface Streamlit
st.title('Gerador de Curva S e Caminho Crítico com Alerta de Atraso e Relatório PDF')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="DD/MM/AAAA")
end_date = st.text_input("Selecione a data final do cronograma (DD/MM/AAAA)", placeholder="DD/MM/AAAA")

if uploaded_file is not None and start_date and end_date:
    try:
        # Converter as datas de início e final para o formato correto
        start_date = pd.to_datetime(start_date, format='%d/%m/%Y')
        end_date = pd.to_datetime(end_date, format='%d/%m/%Y')

        df = read_excel(uploaded_file)
    
        st.write("Dados do cronograma:")
        st.dataframe(df)
    
        atividades_maior_15_dias, atividades_sem_predecessora, caminho_critico = calcular_caminho_critico_maior_que_15_dias(df)

        # Expander para "Atividades sem predecessoras"
        with st.expander("Atividades sem Predecessoras"):
            if atividades_sem_predecessora:
                st.write("Atividades sem predecessoras:")
                atividades_sem_predecessora_df = pd.DataFrame(atividades_sem_predecessora)
                st.table(atividades_sem_predecessora_df[['Nome da tarefa', 'Início', 'Término', 'Duracao']])
            else:
                st.write("Nenhuma atividade sem predecessoras encontrada.")

        # Expander para "Caminho Crítico"
        with st.expander("Caminho Crítico"):
            if atividades_maior_15_dias.empty:
                st.write("Nenhuma atividade com mais de 15 dias de duração no caminho crítico.")
            else:
                st.write("Atividades no caminho crítico com mais de 15 dias de duração:")
                st.table(atividades_maior_15_dias)

        # Gerar alerta de atividades atrasadas
        gerar_alerta_atraso(df)

        if end_date <= start_date:
            st.error("A data final do cronograma deve ser posterior à data inicial.")
        else:
            timeline, curva_s, delta = generate_s_curve(df, start_date, end_date)
            
            st.write("Curva S:")
            curva_s_path = plot_s_curve(timeline, curva_s, start_date)
            
            # Exportar o Excel e fornecer o download
            excel_data = export_to_excel(df, caminho_critico, curva_s, delta, timeline)
            
            # Botão de download do Excel
            st.download_button(
                label="Baixar Cronograma com Curva S",
                data=excel_data.getvalue(),
                file_name="cronograma_com_curva_s.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Gerar relatório em PDF
            atividades_atrasadas = df[df['Término'] < pd.Timestamp.today().normalize()]
            pdf_data = gerar_relatorio_pdf(df, caminho_critico, atividades_sem_predecessora, atividades_atrasadas, curva_s_path)
            
            # Botão de download do PDF
            st.download_button(
                label="Baixar Relatório em PDF",
                data=pdf_data.getvalue(),
                file_name="relatorio_projeto.pdf",
                mime="application/pdf"
            )

    except ValueError:
        st.error("Por favor, insira as datas no formato DD/MM/AAAA.")
