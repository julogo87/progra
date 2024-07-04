from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import os

app = Flask(__name__)

def process_and_plot(df, additional_text):
    try:
        df['fecha_salida'] = pd.to_datetime(df['STD'], format='%d%b %H:%M', dayfirst=True)
        df['fecha_llegada'] = pd.to_datetime(df['STA'], format='%d%b %H:%M', dayfirst=True)
    except KeyError as e:
        return None, f"Missing column in input data: {e}"
    except ValueError as e:
        return None, f"Date conversion error: {e}"

    order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
    df['aeronave'] = pd.Categorical(df['Reg.'], categories=order, ordered=True)
    df = df.sort_values('aeronave', ascending=False)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = f'Programación de Vuelos QT {additional_text}'

    # Escribir la cabecera
    sheet.append(['Aeronave', 'Fecha Salida', 'Fecha Llegada', 'Duración', 'Flight', 'From', 'To'])

    # Escribir los datos
    for _, vuelo in df.iterrows():
        aeronave = vuelo['aeronave']
        start = vuelo['fecha_salida']
        end = vuelo['fecha_llegada']
        duration = end - start
        duration_str = f"{duration.components.hours:02}:{duration.components.minutes:02}"
        sheet.append([aeronave, start.strftime('%d-%b %H:%M'), end.strftime('%d-%b %H:%M'), duration_str, vuelo['Flight'], vuelo['From'], vuelo['To']])

    # Formatear las franjas horarias
    fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=2, max_col=3):
        for cell in row:
            cell.fill = fill

    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)

    return buf, None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        table_data = request.form['table_data']
        additional_text = request.form.get('additional_text')
        try:
            df = pd.read_json(table_data)
        except ValueError as e:
            return jsonify({'error': f"JSON parsing error: {e}"}), 400

        excel, error = process_and_plot(df, additional_text)
        if error:
            return jsonify({'error': error}), 400
        return send_file(excel, as_attachment=True, download_name='programacion_vuelos_qt.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
