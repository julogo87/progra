import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from IPython.display import display
import ipywidgets as widgets
import io
import os
from tkinter import Tk, filedialog

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
    # Aproximación basada en la longitud del texto y la duración de la franja
    text_length_approx = len(text) * 0.02  # Aproximación de la longitud del texto en datos
    return duration.total_seconds() / 3600 >= text_length_approx  # Convertir duración a horas

# Función para procesar los datos y generar el gráfico
def process_and_plot(uploader, additional_text, save_pdf=False, pdf_path=None):
    df = read_uploaded_file(uploader)
    if df is not None:
        # Convertir las columnas de fecha a tipo datetime
        df['fecha_salida'] = pd.to_datetime(df['fecha_salida'])
        df['fecha_llegada'] = pd.to_datetime(df['fecha_llegada'])

        # Ordenar las aeronaves según el orden especificado e invertir el orden
        order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
        df['aeronave'] = pd.Categorical(df['aeronave'], categories=order, ordered=True)
        df = df.sort_values('aeronave', ascending=False)

        # Crear la figura y el eje
        fig, ax = plt.subplots(figsize=(20, 10))

        # Iterar sobre las aeronaves y agregar los vuelos al gráfico
        for i, aeronave in enumerate(reversed(order)):
            vuelos_aeronave = df[df['aeronave'] == aeronave]
            for _, vuelo in vuelos_aeronave.iterrows():
                start = vuelo['fecha_salida']
                duration = vuelo['fecha_llegada'] - vuelo['fecha_salida']
                rect_height = 0.2  # Ajustar el ancho de las franjas
                ax.broken_barh([(start, duration)], (i - rect_height/2, rect_height), facecolors='red')  # Cambio de color a rojo
                # Agregar el número de vuelo en el centro del rectángulo o debajo si no cabe
                flight_text = vuelo['numero_vuelo']
                if text_fits(ax, flight_text, start, duration):
                    ax.text(start + duration / 2, i, flight_text, ha='center', va='center', color='black', fontsize=8)
                else:
                    ax.text(start + duration / 2, i - rect_height, flight_text, ha='center', va='top', color='black', fontsize=8)
                # Agregar el origen al inicio del rectángulo o debajo si no cabe
                origin_text = vuelo['origen']
                if text_fits(ax, origin_text, start, duration):
                    ax.text(start, i + 0.2, origin_text, ha='left', va='center', color='black', fontsize=8)
                else:
                    ax.text(start, i - rect_height, origin_text, ha='left', va='top', color='black', fontsize=8)
                # Agregar el destino al final del rectángulo o debajo si no cabe
                destination_text = vuelo['destino']
                if text_fits(ax, destination_text, start, duration):
                    ax.text(start + duration, i + 0.2, destination_text, ha='right', va='center', color='black', fontsize=8)
                else:
                    ax.text(start + duration, i - rect_height, destination_text, ha='right', va='top', color='black', fontsize=8)
                # Agregar la hora de salida debajo del origen
                ax.text(start, i - 0.2, vuelo['fecha_salida'].strftime('%H:%M'), ha='left', va='center', color='black', fontsize=6)
                # Agregar la hora de llegada debajo del destino
                ax.text(start + duration, i - 0.2, vuelo['fecha_llegada'].strftime('%H:%M'), ha='right', va='center', color='black', fontsize=6)

        # Formato del eje Y
        ax.set_yticks(range(len(order)))
        ax.set_yticklabels(reversed(order))
        ax.set_ylim(-1, len(order))

        # Formato del eje X
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.xticks(rotation=45, fontsize=10)

        # Aumentar el espacio entre las horas en el gráfico
        ax.set_xlim(df['fecha_salida'].min() - pd.Timedelta(hours=1), df['fecha_llegada'].max() + pd.Timedelta(hours=1))

        # Añadir márgenes para aumentar el espacio entre las horas
        plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.15)

        # Etiquetas y título
        plt.xlabel('Hora')
        plt.ylabel('Aeronave')
        plt.title(f'Programación de Vuelos QT {additional_text}')  # Añadir el texto adicional al título

        # Guardar el gráfico como un archivo PDF si se indica
        if save_pdf and pdf_path:
            plt.savefig(pdf_path, format='pdf')

        # Mostrar el gráfico
        plt.show()
    else:
        print("No se ha cargado ningún archivo.")

# Crear el uploader
uploader = upload_file()

# Crear el cuadro de texto para el título adicional
title_text = widgets.Text(
    value='',
    placeholder='Ingrese texto adicional para el título',
    description='Título:',
    disabled=False
)
display(title_text)

# Crear el botón para procesar el archivo
process_button = widgets.Button(description="Procesar archivo")
display(process_button)

# Crear el botón para exportar a PDF
export_button = widgets.Button(description="Exportar a PDF")
display(export_button)

# Asignar la función al botón de procesar
def on_button_clicked(b):
    process_and_plot(uploader, title_text.value)

process_button.on_click(on_button_clicked)

# Asignar la función al botón de exportar
def on_export_button_clicked(b):
    # Usar Tkinter para abrir un cuadro de diálogo de selección de carpeta
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal de Tkinter
    folder_selected = filedialog.askdirectory(title="Seleccionar carpeta para guardar el PDF")
    if folder_selected:
        pdf_path = os.path.join(folder_selected, 'programacion_vuelos_qt.pdf')
        process_and_plot(uploader, title_text.value, save_pdf=True, pdf_path=pdf_path)

export_button.on_click(on_export_button_clicked)


