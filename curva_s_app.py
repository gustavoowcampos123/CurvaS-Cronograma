import pandas as pd
import matplotlib.pyplot as plt
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
    if 'Duração' in df.columns:
        df['Duracao'] = df['Duração'].str.extract('(\d+)').astype(float)
    
    return df

# Função para gerar a Curva S
def gerar_curva_s(df_raw, start_date_str='16/09/2024'):
    # Limpeza e formatação das datas
    df_raw['Início'] = df_raw['Início'].str.replace(r'[A-Za-z]+\s', '', regex=True)
    df_raw['Término'] = df_raw['Término'].str.replace(r'[A-Za-z]+\s', '', regex=True)
    df_raw['Início'] = pd.to_datetime(df_raw['Início'], format='%d/%m/%y', errors='coerce')
    df_raw['Término'] = pd.to_datetime(df_raw['Término'], format='%d/%m/%y', errors='coerce')

    # Definir a linha do tempo semanal
    start_date = pd.to_datetime(start_date_str)
    end_date = df_raw['Término'].max()
    weeks = pd.date_range(start=start_date, end=end_date, freq='W-MON')

    # Inicializar o progresso por semana
    progress_by_week = pd.DataFrame(weeks, columns=['Data'])
    progress_by_week['% Executado'] = 0.0

    # Distribuir o progresso de cada tarefa
    for i, row in df_raw.iterrows():
        if pd.notna(row['Início']) and pd.notna(row['Término']):
            task_weeks = pd.date_range(start=row['Início'], end=row['Término'], freq='W-MON')
            if len(task_weeks) == 0:
                weekly_progress = 1  # Se a tarefa durar menos de uma semana
                week = row['Início']
                progress_by_week.loc[progress_by_week['Data'] == week, '% Executado'] += weekly_progress
            else:
                weekly_progress = 1 / len(task_weeks)  # Progresso linear ao longo das semanas
                for week in task_weeks:
                    progress_by_week.loc[progress_by_week['Data'] == week, '% Executado'] += weekly_progress

    # Calcular o progresso acumulado
    progress_by_week['% Executado Acumulado'] = progress_by_week['% Executado'].cumsum() * 100

    # Normalizar para que o progresso acumulado chegue a 100%
    max_progress = progress_by_week['% Executado Acumulado'].max()
    if max_progress > 0:
        progress_by_week['% Executado Acumulado'] = (progress_by_week['% Executado Acumulado'] / max_progress) * 100

    # Plotar a Curva S
    plt.figure(figsize=(10, 6))
    plt.plot(progress_by_week['Data'], progress_by_week['% Executado Acumulado'], marker='o', linestyle='-', color='b')
    plt.title('Curva S - % Executado por Semana')
    plt.xlabel('Data')
    plt.ylabel('% Executado Acumulado')
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    return progress_by_week

# Função para exportar os dados para Excel com gráfico
def export_to_excel(df, curva_s_df):
    output = io.BytesIO()
    
    wb = Workbook()
    ws = wb.active
    ws.title = 'Curva S'
    
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
    
    wb.save(output)
    output.seek(0)
    
    return output

# Interface Streamlit
st.title('Curva S - Novo Modelo de Execução Semanal')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="16/09/2024")

if st.button("Gerar Curva S"):
    if uploaded_file is not None and start_date:
        try:
            # Carregar o Excel
            df_raw = read_excel(uploaded_file)
            
            # Gerar Curva S
            progress_by_week = gerar_curva_s(df_raw, start_date_str=start_date)
            
            # Exibir o progresso semanal em uma tabela
            st.write(progress_by_week)
            
            # Exportar o Excel e fornecer o download
            excel_data = export_to_excel(df_raw, progress_by_week)
            
            # Botão de download do Excel
            st.download_button(
                label="Baixar Cronograma com Curva S",
                data=excel_data.getvalue(),
                file_name="cronograma_com_curva_s.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as e:
            st.error(f"Erro ao processar os dados: {e}")
