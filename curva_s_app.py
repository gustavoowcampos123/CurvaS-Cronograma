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

# Função para remover prefixos indesejados das predecessoras
def remove_prefix(predecessor):
    prefixes = ['TT', 'TI', 'II']
    for prefix in prefixes:
        if predecessor.startswith(prefix):
            return predecessor[len(prefix):].strip()
    return predecessor.strip()

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

# Função para exportar os dados para Excel com gráfico
def export_to_excel(df, caminho_critico, curva_s, delta, timeline):
    output = io.BytesIO()
    
    wb = Workbook()
    ws = wb.active
    ws.title = 'Curva S'
    
    # Criar um DataFrame para Curva S
    curva_s_df = pd.DataFrame({'Data': timeline, 'Progresso Acumulado (%)': curva_s, 'Delta': delta})
    
    # Adicionar os dados da Curva S à planilha
    for r in dataframe_to_rows(curva_s_df, index=False, header=True):
        ws.append(r)
    
    # Criar o gráfico de linha para a Curva S
    chart = LineChart()
    chart.title = "Curva S - Progresso Acumulado"
    chart.y_axis.title = 'Progresso Acumulado (%)'
    chart.x_axis.title = 'Data'
    
    data = Reference(ws, min_col=2, min_row=2, max_row=len(curva_s_df) + 1, max_col=2)
    chart.add_data(data, titles_from_data=True)
    
    ws.add_chart(chart, "E5")
    
    # Criar outra aba para o cronograma
    cronograma_ws = wb.create_sheet(title="Cronograma")
    for r in dataframe_to_rows(df, index=False, header=True):
        cronograma_ws.append(r)
    
    # Criar outra aba para o caminho crítico
    caminho_critico_ws = wb.create_sheet(title="Caminho Crítico")
    critical_path_df = pd.DataFrame(caminho_critico, columns=['Atividades Caminho Crítico'])
    for r in dataframe_to_rows(critical_path_df, index=False, header=True):
        caminho_critico_ws.append(r)
    
    wb.save(output)
    output.seek(0)
    
    return output

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

# Função para gerar a Curva S
def generate_s_curve(df, start_date, end_date):
    df['Início'] = pd.to_datetime(df['Início'])
    df['Término'] = pd.to_datetime(df['Término'])
    
    df['Duracao'] = (df['Término'] - df['Início']).dt.days
    df['Progresso Diario'] = np.where(df['Duracao'] == 0, 0, 1 / df['Duracao'])
    
    timeline = pd.date_range(start=start_date, end=end_date, freq='W')
    
    progresso_acumulado = []
    for date in timeline:
        progresso_semanal = df.loc[df['Início'] <= date, 'Progresso Diario'].sum()
        progresso_acumulado.append(progresso_semanal)
    
    progresso_acumulado_percentual = np.cumsum(progresso_acumulado)
    progresso_acumulado_percentual = (progresso_acumulado_percentual / progresso_acumulado_percentual[-1]) * 100
    
    delta = np.diff(progresso_acumulado_percentual, prepend=0)
    
    return timeline, progresso_acumulado_percentual, delta

# Função para gerar alerta de atraso
def gerar_alerta_atraso(df):
    data_atual = pd.Timestamp.today().normalize()  # Data de hoje
    atividades_atrasadas = df[df['Término'] < data_atual]

    if not atividades_atrasadas.empty:
        st.warning("Atividades Atrasadas:")
        st.table(atividades_atrasadas[['Nome da tarefa', 'Início', 'Término', 'Duracao']])
    else:
        st.success("Nenhuma atividade atrasada.")

# Função para calcular o caminho crítico e listar as atividades com duração maior que 15 dias
def calcular_caminho_critico_maior_que_15_dias(df):
    caminho_critico, atividades_sem_predecessora = calculate_critical_path(df)

    if not caminho_critico:
        return pd.DataFrame(), atividades_sem_predecessora, caminho_critico

    # Filtrar atividades no caminho crítico com duração superior a 15 dias
    atividades_caminho_critico = df[df['Nome da tarefa'].isin(caminho_critico)]
    atividades_mais_15_dias = atividades_caminho_critico[atividades_caminho_critico['Duracao'] > 15]

    return atividades_mais_15_dias[['Nome da tarefa', 'Duracao', 'Início', 'Término']], atividades_sem_predecessora, caminho_critico

# Função para calcular o caminho crítico
# Função para calcular o caminho crítico
def calculate_critical_path(df):
    G = nx.DiGraph()
    atividades_sem_predecessora = []  # Lista para armazenar as atividades sem predecessora
    
    if 'Predecessoras' in df.columns:
        for i, row in df.iterrows():
            if pd.notna(row['Predecessoras']):
                predecessoras = str(row['Predecessoras']).split(';')
                for pred in predecessoras:
                    pred_clean = remove_prefix(pred.split('-')[0].strip())
                    try:
                        duration = int(row['Duracao'])
                        if pred_clean:
                            G.add_edge(pred_clean, row['Nome da tarefa'], weight=duration)
                    except ValueError:
                        st.error(f"Duração inválida para a tarefa {row['Nome da tarefa']}: {row['Duracao']} (linha {i+1})")
            else:
                # Adiciona as atividades sem predecessora à lista
                atividades_sem_predecessora.append(row)
    else:
        st.error("A coluna 'Predecessoras' não foi encontrada no arquivo.")
    
    if len(G.nodes) == 0:
        st.error("O grafo de atividades está vazio. Verifique as predecessoras e a duração das atividades.")
        return [], atividades_sem_predecessora

    try:
        critical_path = nx.dag_longest_path(G, weight='weight')
        return critical_path, atividades_sem_predecessora
    except Exception as e:
        st.error(f"Erro ao calcular o caminho crítico: {e}")
        return [], atividades_sem_predecessora

# Interface Streamlit
st.title('Gerador de Curva S e Caminho Crítico com Alerta de Atraso e Relatório PDF')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="DD/MM/AAAA")
end_date = st.text_input("Selecione a data final do cronograma (DD/MM/AAAA)", placeholder="DD/MM/AAAA")

if uploaded_file is not None and start_date and end_date:
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
                # Gerar Curva S
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

    except ValueError as e:
        st.error(f"Erro ao processar os dados: {e}")

