import pandas as pd
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import streamlit as st

# Funções anteriores

def read_excel(file):
    df = pd.read_excel(file)
    return df

def calculate_critical_path(df):
    G = nx.DiGraph()
    for i, row in df.iterrows():
        G.add_edge(row['Atividade'], row['Dependencia'], weight=row['Duracao'])
    
    critical_path = nx.dag_longest_path(G, weight='weight')
    return critical_path

def generate_s_curve(df, start_date):
    df['Data Inicio'] = pd.to_datetime(df['Data Inicio'])
    df['Data Fim'] = pd.to_datetime(df['Data Fim'])
    
    df['Duracao'] = (df['Data Fim'] - df['Data Inicio']).dt.days
    df['Progresso Diario'] = 1 / df['Duracao']
    
    timeline = pd.date_range(start=start_date, end=df['Data Fim'].max(), freq='W')
    progresso_acumulado = []
    
    for date in timeline:
        progresso_semanal = df.loc[df['Data Inicio'] <= date, 'Progresso Diario'].sum()
        progresso_acumulado.append(progresso_semanal)
    
    return timeline, np.cumsum(progresso_acumulado)

def export_to_excel(df, caminho_critico, curva_s, timeline, output_path):
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, sheet_name='Cronograma', index=False)
        
        critical_path_df = pd.DataFrame(caminho_critico, columns=['Atividades Caminho Critico'])
        critical_path_df.to_excel(writer, sheet_name='Caminho Critico', index=False)
        
        curva_s_df = pd.DataFrame({'Data': timeline, 'Progresso Acumulado': curva_s})
        curva_s_df.to_excel(writer, sheet_name='Curva S', index=False)

def plot_s_curve(timeline, curva_s):
    fig, ax = plt.subplots()
    ax.plot(timeline, curva_s, marker='o')
    ax.set_title('Curva S')
    ax.set_xlabel('Data')
    ax.set_ylabel('Progresso Acumulado (%)')
    ax.grid(True)
    st.pyplot(fig)

# Interface Streamlit

st.title('Gerador de Curva S e Caminho Crítico')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.date_input("Selecione a data de início do projeto")

if uploaded_file is not None:
    df = read_excel(uploaded_file)
    
    st.write("Dados do cronograma:")
    st.dataframe(df)
    
    caminho_critico = calculate_critical_path(df)
    st.write("Caminho Crítico:")
    st.write(caminho_critico)
    
    timeline, curva_s = generate_s_curve(df, start_date)
    
    st.write("Curva S:")
    plot_s_curve(timeline, curva_s)
    
    if st.button("Exportar Cronograma com Curva S"):
        output_path = 'cronograma_com_curva_s.xlsx'
        export_to_excel(df, caminho_critico, curva_s, timeline, output_path)
        st.success(f"Arquivo exportado com sucesso! Baixe aqui: [Download {output_path}]({output_path})")
