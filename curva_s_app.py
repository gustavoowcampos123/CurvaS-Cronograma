import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference

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
    
    st.write("Colunas encontradas no arquivo:", df.columns)  # Para inspecionar as colunas
    st.write("Datas do cronograma:", df[['Início', 'Término']])  # Verificar datas após conversão
    return df

# Função para remover prefixos indesejados das predecessoras
def remove_prefix(predecessor):
    prefixes = ['TT', 'TI', 'II']
    for prefix in prefixes:
        if predecessor.startswith(prefix):
            return predecessor[len(prefix):].strip()
    return predecessor.strip()

# Função para calcular o caminho crítico com verificações e exibir número da linha das tarefas sem predecessoras
def calculate_critical_path(df):
    G = nx.DiGraph()
    
    # Verificar se a coluna de predecessoras existe
    if 'Predecessoras' in df.columns:
        for i, row in df.iterrows():
            # Verificar se há predecessoras
            if pd.notna(row['Predecessoras']):
                predecessoras = str(row['Predecessoras']).split(';')
                for pred in predecessoras:
                    pred_clean = remove_prefix(pred.split('-')[0].strip())
                    try:
                        # Verificar se a duração não é nula e está no formato correto
                        duration = int(row['Duração'].split()[0])
                        if pred_clean:
                            G.add_edge(pred_clean, row['Nome da tarefa'], weight=duration)
                    except ValueError:
                        st.error(f"Duração inválida para a tarefa {row['Nome da tarefa']}: {row['Duração']} (linha {i+1})")
            else:
                # Exibir número da linha quando não houver predecessoras
                st.warning(f"A tarefa {row['Nome da tarefa']} (linha {i+1}) não tem predecessoras.")
    else:
        st.error("A coluna 'Predecessoras' não foi encontrada no arquivo.")
    
    if len(G.nodes) == 0:
        st.error("O grafo de atividades está vazio. Verifique as predecessoras e a duração das atividades.")
        return []
    
    try:
        critical_path = nx.dag_longest_path(G, weight='weight')
        return critical_path
    except Exception as e:
        st.error(f"Erro ao calcular o caminho crítico: {e}")
        return []

# Função para gerar a Curva S baseada em semanas
def generate_s_curve(df, start_date, end_date):
    df['Início'] = pd.to_datetime(df['Início'])
    df['Término'] = pd.to_datetime(df['Término'])
    
    df['Duracao'] = (df['Término'] - df['Início']).dt.days
    df['Progresso Diario'] = np.where(df['Duracao'] == 0, 0, 1 / df['Duracao'])
    
    # Cria uma timeline semanalmente a partir da data de início até a data final fornecida pelo usuário
    timeline = pd.date_range(start=start_date, end=end_date, freq='W')
    
    progresso_acumulado = []
    for date in timeline:
        progresso_semanal = df.loc[df['Início'] <= date, 'Progresso Diario'].sum()
        progresso_acumulado.append(progresso_semanal)
    
    # Normalizar o progresso acumulado para 0 a 100%
    progresso_acumulado_percentual = np.cumsum(progresso_acumulado)
    progresso_acumulado_percentual = (progresso_acumulado_percentual / progresso_acumulado_percentual[-1]) * 100
    
    # Calcular a diferença semanal (Delta)
    delta = np.diff(progresso_acumulado_percentual, prepend=0)
    
    return timeline, progresso_acumulado_percentual, delta

# Função para exportar os dados para Excel com gráfico na aba "Curva S"
def export_to_excel(df, caminho_critico, curva_s, delta, timeline):
    output = io.BytesIO()
    
    # Criar o workbook e a planilha
    wb = Workbook()
    ws = wb.active
    ws.title = 'Curva S'
    
    # Adicionar os dados de Curva S e Delta na planilha
    curva_s_df = pd.DataFrame({'Data': timeline, 'Progresso Acumulado (%)': curva_s, 'Delta': delta})
    
    for r in dataframe_to_rows(curva_s_df, index=False, header=True):
        ws.append(r)
    
    # Criar o gráfico de linha da Curva S
    chart = LineChart()
    chart.title = "Curva S - Progresso Acumulado"
    chart.y_axis.title = 'Progresso Acumulado (%)'
    chart.x_axis.title = 'Data'
    
    # Referenciar os dados do gráfico
    data = Reference(ws, min_col=2, min_row=1, max_row=len(curva_s_df) + 1, max_col=2)
    chart.add_data(data, titles_from_data=True)
    
    # Colocar o gráfico na planilha
    ws.add_chart(chart, "E5")
    
    # Adicionar dados e gráfico na planilha do cronograma
    cronograma_ws = wb.create_sheet(title="Cronograma")
    for r in dataframe_to_rows(df, index=False, header=True):
        cronograma_ws.append(r)
    
    # Adicionar o caminho crítico na aba
    caminho_critico_ws = wb.create_sheet(title="Caminho Critico")
    critical_path_df = pd.DataFrame(caminho_critico, columns=['Atividades Caminho Critico'])
    for r in dataframe_to_rows(critical_path_df, index=False, header=True):
        caminho_critico_ws.append(r)
    
    # Salvar o Excel no buffer
    wb.save(output)
    output.seek(0)
    
    return output

# Função para plotar a Curva S
def plot_s_curve(timeline, curva_s):
    fig, ax = plt.subplots()
    ax.plot(timeline, curva_s, marker='o', label="Curva S (0 a 100%)")
    
    # Marcar a linha de início do cronograma
    ax.axvline(x=timeline[0], color='green', linestyle='--', label="Início do Cronograma")
    
    # Configurações do gráfico
    ax.set_title('Curva S - Progresso Acumulado (0 a 100%)')
    ax.set_xlabel('Data')
    ax.set_ylabel('Progresso Acumulado (%)')
    ax.set_ylim(0, 100)  # Limitar o eixo Y de 0 a 100%
    ax.grid(True)
    plt.xticks(rotation=45)
    plt.legend()
    
    st.pyplot(fig)

# Interface Streamlit
st.title('Gerador de Curva S e Caminho Crítico')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.date_input("Selecione a data de início do projeto")
end_date = st.date_input("Selecione a data final do cronograma")

if uploaded_file is not None:
    df = read_excel(uploaded_file)
    
    st.write("Dados do cronograma:")
    st.dataframe(df)
    
    caminho_critico = calculate_critical_path(df)
    st.write("Caminho Crítico:")
    st.write(caminho_critico)
    
    if end_date <= start_date:
        st.error("A data final do cronograma deve ser posterior à data inicial.")
    else:
        timeline, curva_s, delta = generate_s_curve(df, start_date, end_date)
        
        st.write
