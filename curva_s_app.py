import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
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
    
    return df

# Função para remover prefixos indesejados das predecessoras
def remove_prefix(predecessor):
    prefixes = ['TT', 'TI', 'II']
    for prefix in prefixes:
        if predecessor.startswith(prefix):
            return predecessor[len(prefix):].strip()
    return predecessor.strip()

# Função para calcular o caminho crítico e exibir atividades sem predecessora em uma tabela
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
    
    # Verificar se há atividades sem predecessoras e exibi-las em uma tabela
    if atividades_sem_predecessora:
        st.write("Atividades sem predecessoras:")
        atividades_sem_predecessora_df = pd.DataFrame(atividades_sem_predecessora)
        st.table(atividades_sem_predecessora_df[['Nome da tarefa', 'Início', 'Término', 'Duracao']])
    
    if len(G.nodes) == 0:
        st.error("O grafo de atividades está vazio. Verifique as predecessoras e a duração das atividades.")
        return []

    try:
        critical_path = nx.dag_longest_path(G, weight='weight')
        return critical_path
    except Exception as e:
        st.error(f"Erro ao calcular o caminho crítico: {e}")
        return []

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

# Função para exportar os dados para Excel com gráfico
def export_to_excel(df, caminho_critico, curva_s, delta, timeline):
    output = io.BytesIO()
    
    wb = Workbook()
    ws = wb.active
    ws.title = 'Curva S'
    
    curva_s_df = pd.DataFrame({'Data': timeline, 'Progresso Acumulado (%)': curva_s, 'Delta': delta})
    
    for r in dataframe_to_rows(curva_s_df, index=False, header=True):
        ws.append(r)
    
    chart = LineChart()
    chart.title = "Curva S - Progresso Acumulado"
    chart.y_axis.title = 'Progresso Acumulado (%)'
    chart.x_axis.title = 'Data'
    
    data = Reference(ws, min_col=2, min_row=2, max_row=len(curva_s_df) + 1, max_col=2)
    chart.add_data(data, titles_from_data=True)
    
    ws.add_chart(chart, "E5")
    
    cronograma_ws = wb.create_sheet(title="Cronograma")
    for r in dataframe_to_rows(df, index=False, header=True):
        cronograma_ws.append(r)
    
    caminho_critico_ws = wb.create_sheet(title="Caminho Critico")
    critical_path_df = pd.DataFrame(caminho_critico, columns=['Atividades Caminho Critico'])
    for r in dataframe_to_rows(critical_path_df, index=False, header=True):
        caminho_critico_ws.append(r)
    
    wb.save(output)
    output.seek(0)
    
    return output

# Função para calcular o caminho crítico e listar as atividades com duração maior que 15 dias
def calcular_caminho_critico_maior_que_15_dias(df):
    caminho_critico = calculate_critical_path(df)

    if not caminho_critico:
        return [], "Caminho crítico não encontrado"

    # Filtrar atividades no caminho crítico com duração superior a 15 dias
    atividades_caminho_critico = df[df['Nome da tarefa'].isin(caminho_critico)]
    atividades_mais_15_dias = atividades_caminho_critico[atividades_caminho_critico['Duracao'] > 15]

    return atividades_mais_15_dias[['Nome da tarefa', 'Duracao', 'Início', 'Término']], caminho_critico

# Função para plotar a Curva S
def plot_s_curve(timeline, curva_s):
    fig, ax = plt.subplots()
    ax.plot(timeline, curva_s, marker='o', label="Curva S (0 a 100%)")
    ax.axvline(x=timeline[0], color='green', linestyle='--', label="Início do Cronograma")
    
    ax.set_title('Curva S - Progresso Acumulado (0 a 100%)')
    ax.set_xlabel('Data')
    ax.set_ylabel('Progresso Acumulado (%)')
    ax.set_ylim(0, 100)
    ax.grid(True)
    plt.xticks(rotation=45)
    plt.legend()
    
    st.pyplot(fig)

# Interface Streamlit
st.title('Gerador de Curva S e Caminho Crítico')

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
    
        caminho_critico = calculate_critical_path(df)
        st.write("Caminho Crítico:")
        st.write(caminho_critico)
        
        # Mostrar as atividades no caminho crítico com duração maior que 15 dias
                # Exibir as atividades com duração superior a 15 dias
        st.write("Atividades no caminho crítico com mais de 15 dias de duração:")
        st.table(atividades_maior_15_dias)

        if end_date <= start_date:
            st.error("A data final do cronograma deve ser posterior à data inicial.")
        else:
            timeline, curva_s, delta = generate_s_curve(df, start_date, end_date)
            
            st.write("Curva S:")
            plot_s_curve(timeline, curva_s)
            
            # Exportar o Excel e fornecer o download
            excel_data = export_to_excel(df, caminho_critico, curva_s, delta, timeline)
            
            # Botão de download
            st.download_button(
                label="Baixar Cronograma com Curva S",
                data=excel_data.getvalue(),
                file_name="cronograma_com_curva_s.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except ValueError:
        st.error("Por favor, insira as datas no formato DD/MM/AAAA.")

