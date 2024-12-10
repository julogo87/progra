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

    # Verificar si todas las aeronaves están presentes en la columna 'Reg.'
    missing_aircraft = [aircraft for aircraft in order if aircraft not in df['Reg.'].unique()]
    if missing_aircraft:
        return None, f"Favor ingresar contenido para las matriculas faltantes: {', '.join(missing_aircraft)}"

    df['aeronave'] = pd.Categorical(df['Reg.'], categories=order, ordered=True)
    df = df.sort_values('aeronave', ascending=False)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Programación de Vuelos QT'

    # Escribir el título y el texto adicional
    sheet.merge_cells('B1:AX1')
    sheet['B1'] = 'PROGRAMACION DE VUELOS Y TRIPULACIONES'
    sheet['B1'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['B1'].font = Font(size=22, bold=True)

    sheet.merge_cells('B2:AX2')
    sheet['B2'] = additional_text
    sheet['B2'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['B2'].font = Font(size=22, italic=True)

    # Formato y estilos
    fill_blue = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    medium_border = Border(left=Side(style='medium'),
                           right=Side(style='medium'),
                           top=Side(style='medium'),
                           bottom=Side(style='medium'))

    start_time = df['fecha_salida'].min().floor('H')
    end_time = df['fecha_llegada'].max().ceil('H')
    num_columns = int((end_time - start_time).total_seconds() / 900) + 1

    # Generar franjas de vuelo y agregar detalles
    for aeronave in order:
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        if vuelos_aeronave.empty:
            continue

        for _, vuelo in vuelos_aeronave.iterrows():
            start = vuelo['fecha_salida']
            end = vuelo['fecha_llegada']
            duration = end - start
            duration_minutes = duration.total_seconds() / 60
            start_col = 2 + int((start - start_time).total_seconds() / 900)
            end_col = start_col + int(duration_minutes / 15)

            current_row = 7  # Actualizar según la lógica de ubicación

            # Agregar franjas de vuelo
            for col in range(start_col, end_col + 1):
                sheet.cell(row=current_row + 1, column=col).fill = fill_blue
                sheet.cell(row=current_row + 2, column=col).fill = fill_blue
                sheet.cell(row=current_row + 3, column=col).fill = fill_yellow

            # Agregar texto
            mid_col = start_col + (end_col - start_col) // 2
            sheet.cell(row=current_row + 2, column=mid_col).value = vuelo['Flight']
            sheet.cell(row=current_row + 2, column=mid_col).alignment = Alignment(horizontal='center', vertical='center')
            sheet.cell(row=current_row + 2, column=mid_col).font = Font(bold=True)

            # Notas debajo del vuelo
            sheet.merge_cells(start_row=current_row + 4, start_column=start_col, end_row=current_row + 4, end_column=end_col)
            sheet.cell(row=current_row + 4, column=start_col).value = vuelo['Notas']
            sheet.cell(row=current_row + 4, column=start_col).alignment = Alignment(horizontal='center', vertical='center')

            # Crew debajo de las notas
            sheet.merge_cells(start_row=current_row + 5, start_column=start_col, end_row=current_row + 5, end_column=end_col)
            sheet.cell(row=current_row + 5, column=start_col).value = vuelo['Crew']
            sheet.cell(row=current_row + 5, column=start_col).alignment = Alignment(horizontal='center', vertical='center')

            # TripADI debajo del Crew
            sheet.merge_cells(start_row=current_row + 6, start_column=start_col, end_row=current_row + 6, end_column=end_col)
            sheet.cell(row=current_row + 6, column=start_col).value = vuelo['TripADI']
            sheet.cell(row=current_row + 6, column=start_col).alignment = Alignment(horizontal='center', vertical='center')

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

