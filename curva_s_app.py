import os
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
    for atividade in caminho_critico['Nome da tarefa']:
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
    pdf.output(output)  # Salvar no buffer de memória
    output.seek(0)
    
    # Verificar se o arquivo temporário ainda existe antes de removê-lo
    if os.path.exists(img_filename):
        os.remove(img_filename)

    return output

# Função para gerar a Curva S
def gerar_curva_s(df_raw, start_date_str='16/09/2024'):
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

# Interface Streamlit
st.title('Gerador de Curva S e Relatório')

uploaded_file = st.file_uploader("Escolha o arquivo Excel do cronograma", type="xlsx")

start_date = st.text_input("Selecione a data de início do projeto (DD/MM/AAAA)", placeholder="16/09/2024")

if st.button("Gerar Relatório"):
    if uploaded_file is not None and start_date:
        try:
            # Carregar o Excel
            df_raw = pd.read_excel(uploaded_file)

            # Geração da Curva S como imagem
            progress_by_week, curva_s_img = gerar_curva_s(df_raw, start_date_str=start_date)

            # Exemplo de outras tabelas
            atividades_sem_predecessora = df_raw[df_raw['Predecessoras'].isna()]
            caminho_critico = df_raw[df_raw['Duracao'] > 15]  # Exemplo simplificado de caminho crítico
            atividades_atrasadas = df_raw[df_raw['Término'] < pd.Timestamp.today()]

            # Exibição nas abas
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Dados do Cronograma", 
                                                          "Atividades sem Predecessora", 
                                                          "Caminho Crítico", 
                                                          "Atividades Atrasadas", 
                                                          "Atividades para Próxima Semana", 
                                                          "Atividades para os Próximos 15 Dias"])
            
            with tab1:
                st.write("### Dados do Cronograma")
                st.dataframe(df_raw)

            with tab2:
                st.write("### Atividades sem Predecessora")
                if not atividades_sem_predecessora.empty:
                    st.dataframe(atividades_sem_predecessora)
                else:
                    st.write("Nenhuma atividade sem predecessora encontrada.")

            with tab3:
                st.write("### Caminho Crítico")
                if not caminho_critico.empty:
                    st.dataframe(caminho_critico)
                else:
                    st.write("Nenhuma atividade com mais de 15 dias de duração no caminho crítico.")

            with tab4:
                st.write("### Atividades Atrasadas")
                if not atividades_atrasadas.empty:
                    st.dataframe(atividades_atrasadas)
                else:
                    st.write("Nenhuma atividade atrasada.")

            proximos_7_dias = pd.Timestamp.today() + pd.Timedelta(days=7)
            atividades_proxima_semana = df_raw[(df_raw['Início'] <= proximos_7_dias) & (df_raw['Término'] >= pd.Timestamp.today())]

            with tab5:
                st.write("### Atividades para Próxima Semana")
                if not atividades_proxima_semana.empty:
                    st.dataframe(atividades_proxima_semana)
                else:
                    st.write("Nenhuma atividade para a próxima semana.")

            proximos_15_dias = pd.Timestamp.today() + pd.Timedelta(days=15)
            atividades_proximos_15_dias = df_raw[(df_raw['Início'] <= proximos_15_dias) & (df_raw['Término'] >= pd.Timestamp.today())]

            with tab6:
                st.write("### Atividades para os Próximos 15 Dias")
                if not atividades_proximos_15_dias.empty:
                    st.dataframe(atividades_proximos_15_dias)
                else:
                    st.write("Nenhuma atividade para os próximos 15 dias.")

            # Exportar os dados para Excel e fornecer o botão de download
            curva_s_df = pd.DataFrame({
                'Data': progress_by_week['Data'],
                '% Executado Acumulado': progress_by_week['% Executado Acumulado']
            })

            excel_data = export_to_excel(df_raw, curva_s_df)

            st.download_button(
                label="Baixar Cronograma com Curva S",
                data=excel_data.getvalue(),
                file_name="cronograma_com_curva_s.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Gerar relatório em PDF
            pdf_data = gerar_relatorio_pdf(df_raw, atividades_sem_predecessora, atividades_atrasadas, caminho_critico, curva_s_img)

            # Botão para baixar o relatório em PDF
            st.download_button(
                label="Baixar Relatório em PDF",
                data=pdf_data.getvalue(),
                file_name="relatorio_projeto.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"Erro ao processar os dados: {e}")

