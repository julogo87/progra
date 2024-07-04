import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from IPython.display import display
import ipywidgets as widgets
import io

# Función para cargar el archivo CSV
def upload_file():
    uploader = widgets.FileUpload(
        accept='.csv',
        multiple=False
    )
    display(uploader)
    return uploader

# Función para leer el archivo cargado
def read_uploaded_file(uploader):
    if uploader.value:
        file_info = next(iter(uploader.value.values()))
        file_content = io.BytesIO(file_info['content'])
        df = pd.read_csv(file_content, sep=';')
        return df
    else:
        return None

# Función para verificar si el texto cabe en la franja
def text_fits(ax, text, start, duration):
    text_length_approx = len(text) * 0.02
    return duration.total_seconds() / 3600 >= text_length_approx

# Función para procesar los datos y generar el gráfico
def process_and_plot(uploader, additional_text):
    df = read_uploaded_file(uploader)
    if df is not None:
        df['fecha_salida'] = pd.to_datetime(df['fecha_salida'])
        df['fecha_llegada'] = pd.to_datetime(df['fecha_llegada'])

        order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
        df['aeronave'] = pd.Categorical(df['aeronave'], categories=order, ordered=True)
        df = df.sort_values('aeronave', ascending=False)

        fig, ax = plt.subplots(figsize=(20, 10))

        for i, aeronave in enumerate(reversed(order)):
            vuelos_aeronave = df[df['aeronave'] == aeronave]
            for _, vuelo in vuelos_aeronave.iterrows():
                start = vuelo['fecha_salida']
                duration = vuelo['fecha_llegada'] - vuelo['fecha_salida']
                rect_height = 0.2
                ax.broken_barh([(start, duration)], (i - rect_height/2, rect_height), facecolors='red')
                flight_text = vuelo['numero_vuelo']
                if text_fits(ax, flight_text, start, duration):
                    ax.text(start + duration / 2, i, flight_text, ha='center', va='center', color='black', fontsize=8)
                else:
                    ax.text(start + duration / 2, i - rect_height, flight_text, ha='center', va='top', color='black', fontsize=8)
                origin_text = vuelo['origen']
                if text_fits(ax, origin_text, start, duration):
                    ax.text(start, i + 0.2, origin_text, ha='left', va='center', color='black', fontsize=8)
                else:
                    ax.text(start, i - rect_height, origin_text, ha='left', va='top', color='black', fontsize=8)
                destination_text = vuelo['destino']
                if text_fits(ax, destination_text, start, duration):
                    ax.text(start + duration, i + 0.2, destination_text, ha='right', va='center', color='black', fontsize=8)
                else:
                    ax.text(start + duration, i - rect_height, destination_text, ha='right', va='top', color='black', fontsize=8)
                ax.text(start, i - 0.2, vuelo['fecha_salida'].strftime('%H:%M'), ha='left', va='center', color='black', fontsize=6)
                ax.text(start + duration, i - 0.2, vuelo['fecha_llegada'].strftime('%H:%M'), ha='right', va='center', color='black', fontsize=6)

        ax.set_yticks(range(len(order)))
        ax.set_yticklabels(reversed(order))
        ax.set_ylim(-1, len(order))

        ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.xticks(rotation=45, fontsize=10)
        ax.set_xlim(df['fecha_salida'].min() - pd.Timedelta(hours=1), df['fecha_llegada'].max() + pd.Timedelta(hours=1))
        plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.15)
        plt.xlabel('Hora')
        plt.ylabel('Aeronave')
        plt.title(f'Programación de Vuelos QT {additional_text}')
        plt.savefig('programacion_vuelos_qt.pdf', format='pdf')
        plt.show()
    else:
        print("No se ha cargado ningún archivo.")

uploader = upload_file()
title_text = widgets.Text(
    value='',
    placeholder='Ingrese texto adicional para el título',
    description='Título:',
    disabled=False
)
display(title_text)
process_button = widgets.Button(description="Procesar archivo")
display(process_button)
export_button = widgets.Button(description="Exportar a PDF")
display(export_button)

def on_button_clicked(b):
    process_and_plot(uploader, title_text.value)

process_button.on_click(on_button_clicked)

def on_export_button_clicked(b):
    process_and_plot(uploader, title_text.value)

export_button.on_click(on_export_button_clicked)
