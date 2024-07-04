from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
import os

app = Flask(__name__)

def text_fits(ax, text, start, duration):
    text_length_approx = len(text) * 0.02
    return duration.total_seconds() / 3600 >= text_length_approx

def process_and_plot(df, additional_text):
    df['fecha_salida'] = pd.to_datetime(df['STD'], format='%Y-%m-%d %H:%M')
    df['fecha_llegada'] = pd.to_datetime(df['STA'], format='%Y-%m-%d %H:%M')
    order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
    df['aeronave'] = pd.Categorical(df['Reg.'], categories=order, ordered=True)
    df = df.sort_values('aeronave', ascending=False)
    fig, ax = plt.subplots(figsize=(20, 10))

    for i, aeronave in enumerate(reversed(order)):
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        for _, vuelo in vuelos_aeronave.iterrows():
            start = vuelo['fecha_salida']
            duration = vuelo['fecha_llegada'] - vuelo['fecha_salida']
            rect_height = 0.2
            ax.broken_barh([(start, duration)], (i - rect_height/2, rect_height), facecolors='red')
            flight_text = vuelo['Flight']
            if text_fits(ax, flight_text, start, duration):
                ax.text(start + duration / 2, i, flight_text, ha='center', va='center', color='black', fontsize=8)
            else:
                ax.text(start + duration / 2, i - rect_height, flight_text, ha='center', va='top', color='black', fontsize=8)
            origin_text = vuelo['From']
            if text_fits(ax, origin_text, start, duration):
                ax.text(start, i + 0.2, origin_text, ha='left', va='center', color='black', fontsize=8)
            else:
                ax.text(start, i - rect_height, origin_text, ha='left', va='top', color='black', fontsize=8)
            destination_text = vuelo['To']
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

    buf = io.BytesIO()
    plt.savefig(buf, format='pdf')
    buf.seek(0)
    plt.close(fig)
    return buf

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        table_text = request.form['table_text']
        additional_text = request.form.get('additional_text')
        df = pd.read_csv(io.StringIO(table_text), sep='\t')
        pdf = process_and_plot(df, additional_text)
        return send_file(pdf, as_attachment=True, download_name='programacion_vuelos_qt.pdf', mimetype='application/pdf')
    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
