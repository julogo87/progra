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

    # Crear el archivo Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Programación de Vuelos QT'

    # Configurar el título y subtítulo
    sheet.merge_cells('B1:AX1')
    sheet['B1'] = 'PROGRAMACION DE VUELOS Y TRIPULACIONES'
    sheet['B1'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['B1'].font = Font(size=22, bold=True)

    sheet.merge_cells('B2:AX2')
    sheet['B2'] = additional_text
    sheet['B2'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['B2'].font = Font(size=18, italic=True)

    # Configurar cabecera con horas completas
    start_time = df['fecha_salida'].min().floor('H')
    end_time = df['fecha_llegada'].max().ceil('H')
    num_columns = int((end_time - start_time).total_seconds() / 900) + 1  # 900 segundos = 15 minutos

    hour_header = [''] * 1 + \
                  [((start_time + pd.Timedelta(minutes=15 * i)).strftime('%H:%M') if (start_time + pd.Timedelta(minutes=15 * i)).minute == 0 else '') for i in range(num_columns)]
    for col in range(len(hour_header)):
        sheet.cell(row=5, column=col + 2).value = hour_header[col]
        sheet.cell(row=5, column=col + 2).font = Font(bold=True, color="8B0000")
        sheet.cell(row=5, column=col + 2).alignment = Alignment(horizontal='center', vertical='center')

    # Ajustar el ancho de las columnas
    for col in range(2, 2 + num_columns):
        sheet.column_dimensions[get_column_letter(col)].width = 2.6

    # Definir formatos y rellenos
    fill_blue = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    fill_gray = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    medium_border = Border(left=Side(style='medium'), right=Side(style='medium'),
                           top=Side(style='medium'), bottom=Side(style='medium'))

    # Escribir las aeronaves en la columna A con formato y bordes
    merge_ranges = [(6, 15), (16, 24), (25, 33), (34, 42), (43, 51), (52, 60), (61, 69)]
    for i, (start_row, end_row) in enumerate(merge_ranges):
        sheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        cell = sheet.cell(row=start_row, column=1)
        cell.value = order[i]
        cell.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
        cell.font = Font(size=28, bold=True)
        for row in range(start_row, end_row + 1):
            sheet.cell(row=row, column=1).border = medium_border

    # Procesar vuelos para cada aeronave
    base_row = 7
    for aeronave in order:
        vuelos = df[df['aeronave'] == aeronave]
        if vuelos.empty:
            continue

        for _, vuelo in vuelos.iterrows():
            start_col = 2 + int((vuelo['fecha_salida'] - start_time).total_seconds() / 900)
            duration = vuelo['fecha_llegada'] - vuelo['fecha_salida']
            end_col = start_col + int(duration.total_seconds() / 900)

            # Dibujar franja de vuelo
            for col in range(start_col, end_col + 1):
                sheet.cell(row=base_row, column=col).fill = fill_blue
                sheet.cell(row=base_row + 1, column=col).fill = fill_yellow

            # Escribir información en la franja
            mid_col = start_col + (end_col - start_col) // 2
            sheet.cell(row=base_row, column=mid_col).value = vuelo['Flight']
            sheet.cell(row=base_row, column=mid_col).font = Font(bold=True)
            sheet.cell(row=base_row, column=mid_col).alignment = Alignment(horizontal='center')

            # Crew, Tripadi y Notas
            sheet.cell(row=base_row + 2, column=mid_col).value = f"Crew: {vuelo.get('Crew', '')}"
            sheet.cell(row=base_row + 3, column=mid_col).value = f"Tripadi: {vuelo.get('Tripadi', '')}"
            sheet.cell(row=base_row + 4, column=mid_col).value = f"Notas: {vuelo.get('Notas', '')}"

        base_row += 10

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
