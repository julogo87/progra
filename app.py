from flask import Flask, request, render_template, send_file
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io

app = Flask(__name__)

def process_and_plot(file_content, additional_text):
    df = pd.read_csv(io.BytesIO(file_content), sep=';')
    
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
            if duration.total_seconds() / 3600 >= len(flight_text) * 0.02:  # Aproximación de la longitud del texto en datos
                ax.text(start + duration / 2, i, flight_text, ha='center', va='center', color='black', fontsize=8)
            else:
                ax.text(start + duration / 2, i - rect_height, flight_text, ha='center', va='top', color='black', fontsize=8)
            # Agregar el origen al inicio del rectángulo o debajo si no cabe
            origin_text = vuelo['origen']
            if duration.total_seconds() / 3600 >= len(origin_text) * 0.02:
                ax.text(start, i + 0.2, origin_text, ha='left', va='center', color='black', fontsize=8)
            else:
                ax.text(start, i - rect_height, origin_text, ha='left', va='top', color='black', fontsize=8)
            # Agregar el destino al final del rectángulo o debajo si no cabe
            destination_text = vuelo['destino']
            if duration.total_seconds() / 3600 >= len(destination_text) * 0.02:
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

    # Guardar el gráfico como un archivo PDF
    pdf_output = io.BytesIO()
    plt.savefig(pdf_output, format='pdf')
    pdf_output.seek(0)

    return pdf_output

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    additional_text = request.form['title_text']
    if file and file.filename.endswith('.csv'):
        pdf_output = process_and_plot(file.read(), additional_text)
        return send_file(pdf_output, as_attachment=True, download_name='programacion_vuelos_qt.pdf', mimetype='application/pdf')
    else:
        return "Por favor, cargue un archivo CSV válido."

if __name__ == "__main__":
    app.run(debug=True)

