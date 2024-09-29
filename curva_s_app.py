import streamlit as st
import requests
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
from geopy.geocoders import Nominatim

# Lista de cidades do estado de São Paulo com coordenadas
coordenadas = {
    # Adicione suas cidades e coordenadas aqui
}

def obter_coordenadas(cidade):
    if cidade in coordenadas:
        return coordenadas[cidade]

    try:
        geolocator = Nominatim(user_agent="geoapiExercises")
        location = geolocator.geocode(cidade)
        if location:
            return location.latitude, location.longitude
    except Exception as e:
        st.error(f"Erro ao obter coordenadas: {e}")

    return None, None

def obter_previsao(cidade):
    lat, lon = obter_coordenadas(cidade)
    
    if lat is None or lon is None:
        return None
    
    url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&daily=precipitation_sum&timezone=America/Sao_Paulo"
    resposta = requests.get(url)
    return resposta.json()

def processar_dados(data):
    previsao = {'data': [], 'precipitação': []}
    
    for item in data['daily']['time']:
        precip = data['daily']['precipitation_sum'][data['daily']['time'].index(item)]
        
        previsao['data'].append(datetime.strptime(item, '%Y-%m-%d').date())
        previsao['precipitação'].append(precip)
        
    return pd.DataFrame(previsao)

def plotar_grafico_polar(df):
    # Configurações do gráfico polar
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw={'projection': 'polar'})

    # Número de dias
    num_dias = len(df)
    theta = np.linspace(0, 2 * np.pi, num_dias, endpoint=False).tolist()
    theta += theta[:1]  # Para fechar o gráfico

    # Dados de precipitação
    precip = df['precipitação'].tolist()
    precip += precip[:1]  # Para fechar o gráfico

    # Criar o gráfico
    ax.fill(theta, precip, color='skyblue', alpha=0.6)
    ax.set_xticks(theta[:-1])  # Marcas dos ângulos
    ax.set_xticklabels(df['data'].dt.day.astype(str), fontsize=12)  # Números dos dias
    ax.set_yticks(np.arange(0, max(precip)+1, 1))  # Marca da precipitação
    ax.set_title('Previsão de Precipitação - Gráfico Polar', va='bottom', fontsize=16)

    return fig

# Interface do Streamlit
st.title('Previsão de Chuva para Obras de Engenharia')
cidade = st.text_input('Digite o nome da cidade (ex: São Paulo, Campinas, etc.):')

if st.button('Obter Previsão'):
    if cidade:
        dados = obter_previsao(cidade)

        if dados and 'daily' in dados:
            df = processar_dados(dados)
            fig = plotar_grafico_polar(df)

            st.pyplot(fig)
        else:
            st.error('Cidade não encontrada ou dados indisponíveis.')
    else:
        st.warning('Por favor, insira o nome da cidade.')