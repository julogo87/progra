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

    # Escribir la cabecera con horas completas en negrita y color vinotinto
    start_time = df['fecha_salida'].min().floor('h')
    end_time = df['fecha_llegada'].max().ceil('h')
    num_columns = int((end_time - start_time).total_seconds() / 900) + 1  # 900 segundos = 15 minutos

    hour_header = [''] * 1 + \
                  [((start_time + pd.Timedelta(minutes=15 * i)).strftime('%H:%M') if (start_time + pd.Timedelta(minutes=15 * i)).minute == 0 else '') for i in range(num_columns)]
    for col in range(len(hour_header)):
        sheet.cell(row=5, column=col + 2).value = hour_header[col]
        sheet.cell(row=5, column=col + 2).font = Font(bold=True, color="8B0000")  # Vinotinto
        sheet.cell(row=5, column=col + 2).alignment = Alignment(horizontal='center', vertical='center')

    # Ajustar el ancho de las columnas desde B
    for col in range(2, 2 + num_columns):
        sheet.column_dimensions[get_column_letter(col)].width = 2.5

    # Formatos y rellenos
    fill_blue = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    fill_light_gray = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    medium_border = Border(left=Side(style='medium'),
                           right=Side(style='medium'),
                           top=Side(style='medium'),
                           bottom=Side(style='medium'))
    dashed_red_border = Border(left=Side(style='dashed', color='FF0000'))

    # Combinar celdas y formato de la columna A
    merge_ranges = [(6, 15), (16, 24), (25, 33), (34, 42), (43, 51), (52, 60), (61, 69)]
    for i, (start_row, end_row) in enumerate(merge_ranges):
        sheet.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        cell = sheet.cell(row=start_row, column=1)
        cell.value = order[i]
        cell.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
        cell.font = Font(size=28, bold=True)

        # Dibujar borde externo grueso en los rangos especificados
        for row in range(start_row, end_row + 1):
            sheet.cell(row=row, column=1).border = medium_border if row in {6, 15, 16, 24, 25, 33, 34, 42, 43, 51, 52, 60, 61, 69} else Border(left=Side(style='thin'))
            sheet.cell(row=row, column=1).fill = fill_light_gray

    # Línea vertical negra a la derecha de las celdas A5 a A69
    for row in range(5, 70):
        cell = sheet.cell(row=row, column=2)
        cell.border = Border(left=Side(style='thick', color='000000'))

    # Agregar líneas horizontales más gruesas
    medium_horizontal_border = Border(top=Side(style='medium'))
    for row in [6, 16, 25, 34, 43, 52, 61, 70]:  # Mover línea de 69 a 70
        for col in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row, column=col)
            cell.border = medium_horizontal_border

    # Agregar líneas verticales rojas en las columnas donde la hora es "05:00"
    for col in range(2, 2 + num_columns):
        if sheet.cell(row=5, column=col).value == "05:00":
            for row in range(5, 70):
                cell = sheet.cell(row=row, column=col - 1)
                cell.border = dashed_red_border

    current_row_offsets = {'N331QT': 0, 'N332QT': -1, 'N334QT': -2, 'N335QT': -3, 'N336QT': -4, 'N337QT': -5}
    base_row = 7  # Iniciar a partir de la fila 7

    for aeronave in order:
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        if vuelos_aeronave.empty:
            continue

        offset = current_row_offsets.get(aeronave, 0)
        current_row = base_row + offset

        for _, vuelo in vuelos_aeronave.iterrows():
            start = vuelo['fecha_salida']
            end = vuelo['fecha_llegada']
            duration = end - start
            duration_minutes = duration.total_seconds() / 60
            start_col = 2 + int((start - start_time).total_seconds() / 900)
            end_col = start_col + int(duration_minutes / 15)

            # Colorear las celdas de la franja horaria
            for col in range(start_col, end_col + 1):
                sheet.cell(row=current_row + 1, column=col).fill = fill_blue
                sheet.cell(row=current_row + 2, column=col).fill = fill_blue
                sheet.cell(row=current_row + 3, column=col).fill = fill_yellow

            # Agregar un recuadro negro alrededor de toda la franja
            for col in range(start_col, end_col + 1):
                if col == start_col:
                    sheet.cell(row=current_row + 1, column=col).border = Border(left=Side(style='medium'), top=Side(style='medium'))
                    sheet.cell(row=current_row + 2, column=col).border = Border(left=Side(style='medium'))
                    sheet.cell(row=current_row + 3, column=col).border = Border(left=Side(style='medium'), bottom=Side(style='medium'))
                elif col == end_col:
                    sheet.cell(row=current_row + 1, column=col).border = Border(right=Side(style='medium'), top=Side(style='medium'))
                    sheet.cell(row=current_row + 2, column=col).border = Border(right=Side(style='medium'))
                    sheet.cell(row=current_row + 3, column=col).border = Border(right=Side(style='medium'), bottom=Side(style='medium'))
                else:
                    sheet.cell(row=current_row + 1, column=col).border = Border(top=Side(style='medium'))
                    sheet.cell(row=current_row + 3, column=col).border = Border(bottom=Side(style='medium'))

            # Colocar el número de vuelo en la celda central de la franja y en negrita
            mid_col = start_col + (end_col - start_col) // 2
            flight_cell = sheet.cell(row=current_row + 2, column=mid_col)
            if not flight_cell.merged:
                flight_cell.value = vuelo['Flight']
                flight_cell.alignment = Alignment(horizontal='center', vertical='center')
                flight_cell.font = Font(bold=True)

            # Colocar el origen y la hora de salida en la primera celda de la franja
            sheet.cell(row=current_row + 1, column=start_col).value = vuelo['From']
            sheet.cell(row=current_row + 2, column=start_col).value = vuelo['fecha_salida'].strftime('%H:%M')

            # Colocar el destino y la hora de llegada dos celdas antes y combinar con las dos siguientes celdas
            destination_range = f"{get_column_letter(end_col - 2)}{current_row + 1}:{get_column_letter(end_col)}{current_row + 1}"
            arrival_range = f"{get_column_letter(end_col - 2)}{current_row + 2}:{get_column_letter(end_col)}{current_row + 2}"

            if destination_range not in sheet.merged_cells:
                sheet.cell(row=current_row + 1, column=end_col - 2).value = vuelo['To']
                sheet.cell(row=current_row + 1, column=end_col - 2).alignment = Alignment(horizontal='right')
                sheet.merge_cells(destination_range)

            if arrival_range not in sheet.merged_cells:
                sheet.cell(row=current_row + 2, column=end_col - 2).value = vuelo['fecha_llegada'].strftime('%H:%M')
                sheet.cell(row=current_row + 2, column=end_col - 2).alignment = Alignment(horizontal='right')
                sheet.merge_cells(arrival_range)

            # Crear una celda combinada debajo de la franja
            combined_range = f"{get_column_letter(start_col)}{current_row + 4}:{get_column_letter(end_col)}{current_row + 4}"
            if combined_range not in sheet.merged_cells:
                sheet.merge_cells(combined_range)

        sheet.append([''] * (1 + num_columns))
        sheet.append([''] * (1 + num_columns))
        current_row += 10  # Mover a la siguiente fila base

    # Combinar celdas A70 a A73 y aplicar formato
    sheet.merge_cells('A70:A73')
    sheet['A70'].value = 'CORR.'
    sheet['A70'].alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
    sheet['A70'].font = Font(size=20, bold=True)
    sheet['A70'].fill = fill_light_gray
    sheet['A70'].border = medium_border

    # Combinar celdas B70 a CW73 y aplicar formato
    sheet.merge_cells('B70:CW73')
    cell = sheet['B70']
    cell.fill = fill_light_gray
    cell.border = medium_border

    # Combinar celdas CX70 a EA73 y aplicar formato
    sheet.merge_cells('CX70:EA73')
    cell = sheet['CX70']
    cell.border = medium_border

    # Combinar celdas A74 a A92 y aplicar formato
    sheet.merge_cells('A74:A92')
    sheet['A74'].value = 'PROGRAMACION TRIPULACIONES'
    sheet['A74'].alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
    sheet['A74'].font = Font(size=20, bold=True)
    sheet['A74'].fill = fill_light_gray
    sheet['A74'].border = medium_border

    # Combinar celdas B74 a W75 y aplicar formato
    sheet.merge_cells('B74:W75')
    cell = sheet['B74']
    cell.value = 'MIAMI'
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.font = Font(size=20, bold=True)
    cell.fill = fill_light_gray
    cell.border = medium_border

    # Combinar celdas X74 a AS75 y aplicar formato
    sheet.merge_cells('X74:AS75')
    cell = sheet['X74']
    cell.value = 'BASES'
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.font = Font(size=20, bold=True)
    cell.fill = fill_light_gray
    cell.border = medium_border

    # Combinar celdas AT74 a BO75 y aplicar formato
    sheet.merge_cells('AT74:BO75')
    cell = sheet['AT74']
    cell.value = 'TRASLADOS'
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.font = Font(size=20, bold=True)
    cell.fill = fill_light_gray
    cell.border = medium_border

    # Combinar celdas BP74 a CW75 y aplicar formato
    sheet.merge_cells('BP74:CW75')
    cell = sheet['BP74']
    cell.value = 'ENTRENAMIENTO'
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.font = Font(size=20, bold=True)
    cell.fill = fill_light_gray
    cell.border = medium_border

    # Combinar celdas CX74 a EA92 y aplicar formato
    sheet.merge_cells('CX74:EA92')
    cell = sheet['CX74']
    cell.border = medium_border

    # Seleccionar celdas B76 a W92 y aplicar formato
    for row in range(76, 93):
        for col in range(2, 24):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(fill_type=None)  # Fondo blanco
            cell.border = medium_border

    # Seleccionar celdas X76 a AS92 y aplicar formato
    for row in range(76, 93):
        for col in range(24, 45):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(fill_type=None)  # Fondo blanco
            cell.border = medium_border

    # Seleccionar celdas AT76 a BO92 y aplicar formato
    for row in range(76, 93):
        for col in range(45, 65):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(fill_type=None)  # Fondo blanco
            cell.border = medium_border

    # Seleccionar celdas BP76 a CW92 y aplicar formato
    for row in range(76, 93):
        for col in range(65, 96):
            cell = sheet.cell(row=row, column=col)
            cell.fill = PatternFill(fill_type=None)  # Fondo blanco
            cell.border = medium_border

    # Configurar el zoom del PDF al 65%
    sheet.sheet_view.zoomScale = 65

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
