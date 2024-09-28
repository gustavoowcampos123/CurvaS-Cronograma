import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import streamlit as st

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
    
    st.write("Colunas encontradas no arquivo:", df.columns)  # Para inspecionar as colunas
    return df

# Função para calcular o caminho crítico
def calculate_critical_path(df):
    G = nx.DiGraph()
    
    # Verificar se a coluna de dependência existe
    if 'Dependencia' in df.columns:
        for i, row in df.iterrows():
            G.add_edge(row['Nome da tarefa'], row['Dependencia'], weight=row['Duração'])
    else:
        st.error("A coluna 'Dependencia' não foi encontrada no arquivo.")
    
    critical_path = nx.dag_longest_path(G, weight='weight')
    return critical_path

# Função para gerar a Curva S baseada em semanas
def generate_s_curve(df, start_date, end_date):
    df['Início'] = pd.to_datetime(df['Início'])
    df['Término'] = pd.to_datetime(df['Término'])
    
    df['Duracao'] = (df['Término'] - df['Início']).dt.days
    df['Progresso Diario'] = 1 / df['Duracao']
    
    # Cria uma timeline semanalmente a partir da data de início até a data final fornecida pelo usuário
    timeline = pd.date_range(start=start_date, end=end_date, freq='W')
    
    progresso_acumulado = []
    
    # Acumula o progresso semanal
    for date in timeline:
        progresso_semanal = df.loc[df['Início'] <= date, 'Progresso Diario'].sum()
        progresso_acumulado.append(progresso_semanal)
    
    return timeline, np.cumsum(progresso_acumulado)

# Função para exportar os dados para Excel
def export_to_excel(df, caminho_critico, curva_s, timeline, output_path):
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, sheet_name='Cronograma', index=False)
        
        critical_path_df = pd.DataFrame(caminho_critico, columns=['Atividades Caminho Critico'])
        critical_path_df.to_excel(writer, sheet_name='Caminho Critico', index=False)
        
        curva_s_df = pd.DataFrame({'Data': timeline, 'Progresso Acumulado': curva_s})
        curva_s_df.to_excel(writer, sheet_name='Curva S', index=False)

# Função para plotar a Curva S
def plot_s_curve(timeline, curva_s):
    fig, ax = plt.subplots()
    ax.plot(timeline, curva_s, marker='o', label="Curva S")
    
    # Marcar a linha de início do cronograma
    ax.axvline(x=timeline[0], color='green', linestyle='--', label="Início do Cronograma")
    
    ax.set_title('Curva S - Progresso Acumulado Semanal')
    ax.set_xlabel('Data')
    ax.set_ylabel('Progresso Acumulado (%)')
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
    
    if end_date and start_date:
        timeline, curva_s = generate_s_curve(df, start_date, end_date)
        
        st.write("Curva S:")
        plot_s_curve(timeline, curva_s)
        
        if st.button("Exportar Cronograma com Curva S"):
            output_path = 'cronograma_com_curva_s.xlsx'
            export_to_excel(df, caminho_critico, curva_s, timeline, output_path)
            st.success(f"Arquivo exportado com sucesso! Baixe aqui: [Download {output_path}]({output_path})")

