import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Título de la aplicación
st.title('Programación de Vuelos QT')

# Subir el archivo CSV
uploaded_file = st.file_uploader("Elige un archivo CSV", type="csv")

if uploaded_file is not None:
    # Cargar los datos desde el archivo CSV con el separador adecuado
    df = pd.read_csv(uploaded_file, sep=';')

    # Convertir las columnas de fecha a tipo datetime
    df['fecha_salida'] = pd.to_datetime(df['fecha_salida'])
    df['fecha_llegada'] = pd.to_datetime(df['fecha_llegada'])

    # Crear la figura y el eje
    fig, ax = plt.subplots(figsize=(15, 10))

    # Iterar sobre las aeronaves y agregar los vuelos al gráfico
    aeronaves = df['aeronave'].unique()
    for i, aeronave in enumerate(aeronaves):
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        for _, vuelo in vuelos_aeronave.iterrows():
            ax.broken_barh([(vuelo['fecha_salida'], vuelo['fecha_llegada'] - vuelo['fecha_salida'])], 
                           (i - 0.4, 0.8), facecolors='red')
            # Agregar el número de vuelo en el centro del rectángulo
            ax.text(vuelo['fecha_salida'] + (vuelo['fecha_llegada'] - vuelo['fecha_salida']) / 2, 
                    i, vuelo['numero_vuelo'], ha='center', va='center', color='white')
            # Agregar el origen al inicio del rectángulo
            ax.text(vuelo['fecha_salida'], i + 0.2, vuelo['origen'], ha='left', va='center', color='white')
            # Agregar el destino al final del rectángulo
            ax.text(vuelo['fecha_llegada'], i + 0.2, vuelo['destino'], ha='right', va='center', color='white')
            # Agregar la hora de salida debajo del origen
            ax.text(vuelo['fecha_salida'], i - 0.2, vuelo['fecha_salida'].strftime('%H:%M'), ha='left', va='center', color='white')
            # Agregar la hora de llegada debajo del destino
            ax.text(vuelo['fecha_llegada'], i - 0.2, vuelo['fecha_llegada'].strftime('%H:%M'), ha='right', va='center', color='white')

    # Formato del eje Y
    ax.set_yticks(range(len(aeronaves)))
    ax.set_yticklabels(aeronaves)
    ax.set_ylim(-1, len(aeronaves))

    # Formato del eje X
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    plt.xticks(rotation=45)

    # Etiquetas y título
    plt.xlabel('Hora')
    plt.ylabel('Aeronave')
    plt.title('Programación de Vuelos QT')

    # Mostrar el gráfico en Streamlit
    st.pyplot(fig)
