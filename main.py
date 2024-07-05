from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
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

    # Escribir el título y el texto adicional
    sheet.merge_cells('H1:Z1')
    sheet['H1'] = 'PROGRAMACION DE VUELOS Y TRIPULACIONES'
    sheet['H1'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['H1'].font = Font(size=14, bold=True)

    sheet.merge_cells('H2:Z2')
    sheet['H2'] = additional_text
    sheet['H2'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['H2'].font = Font(size=12, italic=True)

    # Escribir la cabecera
    start_time = df['fecha_salida'].min().floor('H')
    end_time = df['fecha_llegada'].max().ceil('H')
    num_columns = int((end_time - start_time).total_seconds() / 900) + 1  # 900 segundos = 15 minutos

    header = ['Aeronave', 'Fecha Salida', 'Fecha Llegada', 'Duración', 'Flight', 'From', 'To'] + \
             [(start_time + pd.Timedelta(minutes=15 * i)).strftime('%H:%M') for i in range(num_columns)]
    hour_header = ['Aeronave', 'Fecha Salida', 'Fecha Llegada', 'Duración', 'Flight', 'From', 'To'] + \
                  [((start_time + pd.Timedelta(minutes=15 * i)).strftime('%H:%M') if (start_time + pd.Timedelta(minutes=15 * i)).minute == 0 else '') for i in range(num_columns)]

    # Escribir las horas en la fila 3
    for col in range(len(hour_header)):
        sheet.cell(row=3, column=col + 8).value = hour_header[col]

    # Ajustar el ancho de las columnas
    for col in range(8, 8 + num_columns):
        sheet.column_dimensions[get_column_letter(col)].width = 4.3  # 32 píxeles aproximadamente

    # Eliminar las columnas B a G
    for col in range(2, 8):
        sheet.delete_cols(2)

    # Formatos y rellenos
    fill_blue = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    # Combinar celdas y formato de la columna A
    merge_ranges = [(6, 15), (16, 24), (25, 33), (34, 42), (43, 51), (52, 60), (61, 69)]
    for i, (start_row, end_row) in enumerate(merge_ranges):
        sheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        cell = sheet.cell(row=start_row, column=1)
        cell.value = order[i]
        cell.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
        cell.font = Font(size=28, bold=True)

    # Agregar líneas horizontales
    for row in [6, 16, 25, 34, 43, 52, 61, 69]:
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            cell.border = Border(top=Side(style='thin'))

    current_row = 7  # Iniciar a partir de la fila 7

    for aeronave in order:
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        if vuelos_aeronave.empty:
            continue

        row_data = [aeronave] + [''] * (7 + num_columns)
        sheet.append([''] * (7 + num_columns))
        sheet.append([''] * (7 + num_columns))
        sheet.append(row_data)

        for _, vuelo in vuelos_aeronave.iterrows():
            start = vuelo['fecha_salida']
            end = vuelo['fecha_llegada']
            duration = end - start
            duration_minutes = duration.total_seconds() / 60
            start_col = 8 + int((start - start_time).total_seconds() / 900)
            end_col = start_col + int(duration_minutes / 15)

            # Colorear las celdas de la franja horaria
            for col in range(start_col, end_col + 1):
                sheet.cell(row=current_row + 1, column=col).fill = fill_blue
                sheet.cell(row=current_row + 2, column=col).fill = fill_blue
                sheet.cell(row=current_row + 3, column=col).fill = fill_yellow

            # Colocar el número de vuelo en la celda central de la franja
            mid_col = start_col + (end_col - start_col) // 2
            sheet.cell(row=current_row + 2, column=mid_col).value = vuelo['Flight']
            sheet.cell(row=current_row + 2, column=mid_col).alignment = Alignment(horizontal='center', vertical='center')

            # Colocar el origen y la hora de salida en la primera celda de la franja
            sheet.cell(row=current_row + 1, column=start_col).value = vuelo['From']
            sheet.cell(row=current_row + 2, column=start_col).value = vuelo['fecha_salida'].strftime('%H:%M')

            # Colocar el destino y la hora de llegada en la celda superior derecha de la franja
            sheet.cell(row=current_row + 1, column=end_col).value = vuelo['To']
            sheet.cell(row=current_row + 2, column=end_col).value = vuelo['fecha_llegada'].strftime('%H:%M')

        sheet.append([''] * (7 + num_columns))
        sheet.append([''] * (7 + num_columns))
        current_row += 9

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
