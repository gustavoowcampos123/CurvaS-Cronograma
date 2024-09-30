import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from fpdf import FPDF
from PIL import Image
import datetime
import tempfile  # Para gerar arquivos temporários

# Função para limpar a abreviação dos dias da semana
def clean_weekday_abbreviation(date_str):
    return date_str.split(' ', 1)[1] if isinstance(date_str, str) else date_str

# Função para ler o arquivo Excel e tratar as colunas de data
def read_excel(file):
    df = pd.read_excel(file)
    
    # Limpar as colunas de datas
    df['Início'] = df['Início'].apply(lambda x: clean_weekday_abbreviation(x) if isinstance(x, str) else x)
    df['Término'] = df['Término'].apply(lambda x: clean_weekday_abbreviation(x) if isinstance(x, str) else x)
    
    # Converter para datetime
    df['Início'] = pd.to_datetime(df['Início'], format='%d/%m/%y', errors='coerce')
    df['Término'] = pd.to_datetime(df['Término'], format='%d/%m/%y', errors='coerce')
    
    # Tratar a duração (remover "dias" e converter para float)
    if 'Duração' in df.columns:
        df['Duracao'] = df['Duração'].str.extract('(\d+)').astype(float)
    
    return df

# Função para gerar a Curva S
def gerar_curva_s(df_raw, start_date_str='16/09/2024'):
    df_raw['Início'] = df_raw['Início'].apply(lambda x: clean_weekday_abbreviation(x) if isinstance(x, str) else x)
    df_raw['Término'] = df_raw['Término'].apply(lambda x: clean_weekday_abbreviation(x) if isinstance(x, str) else x)
    
    start_date = pd.to_datetime(start_date_str)
    end_date = df_raw['Término'].max()
    weeks = pd.date_range(start=start_date, end=end_date, freq='W-MON')

    progress_by_week = pd.DataFrame(weeks, columns=['Data'])
    progress_by_week['% Executado'] = 0.0

    for i, row in df_raw.iterrows():
        if pd.notna(row['Início']) and pd.notna(row['Término']):
            task_weeks = pd.date_range(start=row['Início'], end=row['Término'], freq='W-MON')
            if len(task_weeks) == 0:
                weekly_progress = 1  # Se a tarefa durar menos de uma semana
                week = row['Início']
                progress_by_week.loc[progress_by_week['Data'] == week, '% Executado'] += weekly_progress
            else:
                weekly_progress = 1 / len(task_weeks)
                for week in task_weeks:
                    progress_by_week.loc[progress_by_week['Data'] == week, '% Executado'] += weekly_progress

    progress_by_week['% Executado Acumulado'] = progress_by_week['% Executado'].cumsum() * 100

    max_progress = progress_by_week['% Executado Acumulado'].max()
    if max_progress > 0:
        progress_by_week['% Executado Acumulado'] = (progress_by_week['% Executado Acumulado'] / max_progress) * 100

    # Plotar a Curva S e retornar o gráfico como imagem
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(progress_by_week['Data'], progress_by_week['% Executado Acumulado'], marker='o', linestyle='-', color='b')
    ax.set_title('Curva S - % Executado por Semana')
    ax.set_xlabel('Data')
    ax.set_ylabel('% Executado Acumulado')
    ax.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()

    # Converter o gráfico em imagem no formato PNG
    img_buffer = io.BytesIO()
    fig.savefig(img_buffer, format='png')
    img_buffer.seek(0)
    plt.close(fig)

    st.pyplot(fig)

    return progress_by_week, img_buffer

# Função para exportar os dados para Excel com gráfico
def export_to_excel(df, curva_s_df):
    output = io.BytesIO()
    
    wb = Workbook()
    ws = wb.active
    ws.title = 'Curva S'
    
    for r in dataframe_to_rows(curva_s_df, index=False, header=True):
        ws.append(r)
    
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
    
    # Remover o arquivo temporário
    os.remove(img_filename)

    return output

# Interface Streamlit
st.title('Gerador de Curva S e Relatório')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="16/09/2024")

if st.button("Gerar Relatório"):
    if uploaded_file is not None and start_date:
        try:
            # Carregar o Excel
            df_raw = read_excel(uploaded_file)

            # Gerar Curva S e obter o gráfico como imagem
            progress_by_week, curva_s_img = gerar_curva_s(df_raw, start_date_str=start_date)

            # Abas para visualização
            st.write("### Dados do Cronograma")
            st.dataframe(df_raw)

            atividades_sem_predecessora = df_raw[df_raw['Predecessoras'].isna()]

            st.write("### Atividades sem Predecessoras")
            st.dataframe(atividades_sem_predecessora)

            caminho_critico = df_raw[df_raw['Duracao'] > 15]  # Exemplo de caminho crítico simplificado
            st.write("### Caminho Crítico")
            st.dataframe(caminho_critico)

            atividades_atrasadas = df_raw[df_raw['Término'] < pd.Timestamp.today()]
            st.write("### Atividades Atrasadas")
            st.dataframe(atividades_atrasadas)

            proximos_7_dias = pd.Timestamp.today() + pd.Timedelta(days=7)
                        # Atividades para Próxima Semana
            st.write("### Atividades para Próxima Semana")
            st.dataframe(atividades_proxima_semana)

            proximos_15_dias = pd.Timestamp.today() + pd.Timedelta(days=15)
            atividades_proximos_15_dias = df_raw[(df_raw['Início'] <= proximos_15_dias) & (df_raw['Término'] >= pd.Timestamp.today())]
            
            # Atividades para os Próximos 15 dias
            st.write("### Atividades para os Próximos 15 Dias")
            st.dataframe(atividades_proximos_15_dias)

            # Gerar Relatório em PDF com a imagem da Curva S
            pdf_data = gerar_relatorio_pdf(df_raw, atividades_sem_predecessora, atividades_atrasadas, caminho_critico, curva_s_img)

            # Botão para baixar o relatório em PDF
            st.download_button(
                label="Baixar Relatório em PDF",
                data=pdf_data.getvalue(),
                file_name="relatorio_projeto.pdf",
                mime="application/pdf"
            )

        except ValueError as e:
            st.error(f"Erro ao processar os dados: {e}")

