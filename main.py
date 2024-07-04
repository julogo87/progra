from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
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

    df = df.dropna(subset=['fecha_salida', 'fecha_llegada'])
    order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
    df['aeronave'] = pd.Categorical(df['Reg.'], categories=order, ordered=True)
    df = df.sort_values('aeronave', ascending=False)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = f'Programación de Vuelos QT {additional_text}'

    # Escribir la cabecera
    start_time = df['fecha_salida'].min().floor('H')
    end_time = df['fecha_llegada'].max().ceil('H')
    num_columns = int((end_time - start_time).total_seconds() / 900) + 1  # 900 segundos = 15 minutos

    header = ['Aeronave', 'Fecha Salida', 'Fecha Llegada', 'Duración', 'Flight', 'From', 'To'] + \
             [(start_time + pd.Timedelta(minutes=15 * i)).strftime('%H:%M') for i in range(num_columns)]
    sheet.append(header)

    fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    for _, vuelo in df.iterrows():
        aeronave = vuelo['aeronave']
        start = vuelo['fecha_salida']
        end = vuelo['fecha_llegada']
        duration = end - start
        duration_minutes = duration.total_seconds() / 60
        duration_str = f"{int(duration_minutes // 60):02}:{int(duration_minutes % 60):02}"
        row_data = [aeronave, start.strftime('%d-%b %H:%M'), end.strftime('%d-%b %H:%M'), duration_str, vuelo['Flight'], vuelo['From'], vuelo['To']] + \
                   [''] * num_columns
        row_idx = sheet.max_row + 1
        sheet.append(row_data)

        start_col = 8 + int((start - start_time).total_seconds() / 900)
        end_col = start_col + int(duration_minutes / 15)

        for col in range(start_col, end_col + 1):
            sheet.cell(row=row_idx, column=col).fill = fill

        mid_col = start_col + (end_col - start_col) // 2
        sheet.cell(row=row_idx, column=mid_col).value = vuelo['Flight']
        sheet.cell(row=row_idx, column=mid_col).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

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
