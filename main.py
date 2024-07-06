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
    start_time = df['fecha_salida'].min().floor('H')
    end_time = df['fecha_llegada'].max().ceil('H')
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
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
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
                sheet.cell(row=row, column=col - 1).border = dashed_red_border

    current_row_offsets = {'N331QT': 0, 'N332QT': -1, 'N334QT': -2, 'N335QT': -3, 'N336QT': -4, 'N337QT': -5}
    base_row = 7  # Iniciar a partir de la fila 7

    for aeronave in order:
        vuelos_aeronave = df[df['aeronave'] == aeronave]
        if vuelos_aeronave.empty:
            continue

        offset = current_row_offsets.get(aeronave, 0)
        current_row = base_row + offset

        row_data = [''] * (1 + num_columns)
        sheet.append([''] * (1 + num_columns))
        sheet.append([''] * (1 + num_columns))
        sheet.append(row_data)

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
                sheet.cell(row[current_row + 3, column=col).fill = fill_yellow

            # Agregar un recuadro negro alrededor de toda la franja
            for col in range(start_col, end_col + 1):
                if col == start_col:
                    sheet.cell(row=current_row + 1, column=col).border = Border(left=Side(style='medium'), top=Side(style='medium'))
                    sheet.cell(row=current_row + 2, column=col).border = Border(left=Side(style='medium'))
                    sheet.cell[row=current_row + 3, column=col).border = Border(left=Side(style='medium'), bottom=Side(style='medium'))
                elif col == end_col:
                    sheet.cell[row=current_row + 1, column=col).border = Border(right=Side(style='medium'), top=Side(style='medium'))
                    sheet.cell[row=current_row + 2, column=col).border = Border(right=Side(style='medium'))
                    sheet.cell[row=current_row + 3, column=col).border = Border(right=Side(style='medium'), bottom=Side(style='medium'))
                else:
                    sheet.cell[row=current_row + 1, column=col).border = Border(top=Side(style='medium'))
                    sheet.cell[row=current_row + 3, column=col).border = Border(bottom=Side(style='medium'))

            # Colocar el número de vuelo en la celda central de la franja y en negrita
            mid_col = start_col + (end_col - start_col) // 2
            sheet.cell[row=current_row + 2, column=mid_col].value = vuelo['Flight']
            sheet.cell[row=current_row + 2, column=mid_col].alignment = Alignment(horizontal='center', vertical='center')
            sheet.cell[row=current_row + 2, column=mid_col].font = Font(bold=True)

            # Colocar el origen y la hora de salida en la primera celda de la franja
            sheet.cell[row=current_row + 1, column=start_col].value = vuelo['From']
            sheet.cell[row=current_row + 2, column=start_col].value = vuelo['fecha_salida'].strftime('%H:%M')

            # Colocar el destino y la hora de llegada una celda antes y combinar con la siguiente celda
            sheet.cell[row=current_row + 1, column=end_col - 1].value = vuelo['To']
            sheet.cell[row=current_row + 1, column=end_col - 1].alignment = Alignment(horizontal='right')
            sheet.merge_cells(start_row=current_row + 1, start_column=end_col - 1, end_row=current_row + 1, end_column=end_col)

            sheet.cell[row=current_row + 2, column=end_col - 1].value = vuelo['fecha_llegada'].strftime('%H:%M')
            sheet.cell[row=current_row + 2, column=end_col - 1].alignment = Alignment(horizontal='right')
            sheet.merge_cells(start_row=current_row + 2, start_column=end_col - 1, end_row=current_row + 2, end_column=end_col)

            # Crear una celda combinada debajo de la franja
            sheet.merge_cells(start_row=current_row + 4, start_column=start_col, end_row=current_row + 4, end_column=end_col)

        sheet.append([''] * (1 + num_columns))
        sheet.append([''] * (1 + num_columns))
        base_row += 10  # Mover a la siguiente fila base

    # Agregar las nuevas combinaciones de celdas y formateos según las especificaciones
    # CORR.
    sheet.merge_cells('A70:A73')
    cell = sheet['A70']
    cell.value = "CORR."
    cell.fill = fill_light_gray
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
    cell.border = medium_border

    sheet.merge_cells('B70:CW73')
    for row in range(70, 74):
        for col in range(2, 102):
            sheet.cell(row=row, column=col).border = medium_border
            sheet.cell(row=row, column=col).fill = fill_light_gray

    sheet.merge_cells('CX70:EA73')
    for row in range(70, 74):
        for col in range(102, 106):
            sheet.cell(row=row, column=col).border = medium_border

    # PROGRAMACION TRIPULACIONES
    sheet.merge_cells('A74:A92')
    cell = sheet['A74']
    cell.value = "PROGRAMACION TRIPULACIONES"
    cell.fill = fill_light_gray
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center', text_rotation=90)
    cell.border = medium_border

    # MIAMI
    sheet.merge_cells('B74:W75')
    for row in range(74, 76):
        for col in range(2, 24):
            sheet.cell(row=row, column=col).border = medium_border
            sheet.cell(row=row, column=col).fill = fill_light_gray
    cell = sheet['B74']
    cell.value = "MIAMI"
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # BASES
    sheet.merge_cells('X74:AS75')
    for row in range(74, 76):
        for col in range(24, 46):
            sheet.cell(row=row, column=col).border = medium_border
            sheet.cell(row=row, column=col).fill = fill_light_gray
    cell = sheet['X74']
    cell.value = "BASES"
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # TRASLADOS
    sheet.merge_cells('AT74:BO75')
    for row in range(74, 76):
        for col in range(46, 64):
            sheet.cell(row=row, column=col).border = medium_border
            sheet.cell(row=row, column=col).fill = fill_light_gray
    cell = sheet['AT74']
    cell.value = "TRASLADOS"
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # ENTRENAMIENTO
    sheet.merge_cells('BP74:CW75')
    for row in range(74, 76):
        for col in range(64, 102):
            sheet.cell(row=row, column=col).border = medium_border
            sheet.cell(row=row, column=col).fill = fill_light_gray
    cell = sheet['BP74']
    cell.value = "ENTRENAMIENTO"
    cell.font = Font(size=20, bold=True)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Resto de combinaciones y formatos
    sheet.merge_cells('CX74:EA92')
    for row in range(74, 93):
        for col in range(102, 106):
            sheet.cell(row=row, column=col).border = medium_border

    # Selección de celdas con borde exterior grueso y fondo blanco
    def set_outer_border(sheet, start_row, end_row, start_col, end_col, border, fill):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell.fill = fill
                if row == start_row:
                    cell.border = Border(top=border.top) if border.top else cell.border
                if row == end_row:
                    cell.border = Border(bottom=border.bottom) if border.bottom else cell.border
                if col == start_col:
                    cell.border = Border(left=border.left) if border.left else cell.border
                if col == end_col:
                    cell.border = Border(right=border.right) if border.right else cell.border

    set_outer_border(sheet, 76, 92, 2, 23, medium_border, fill_white)  # B76:W92
    set_outer_border(sheet, 76, 92, 24, 45, medium_border, fill_white) # X76:AS92
    set_outer_border(sheet, 76, 92, 46, 63, medium_border, fill_white) # AT76:BO92
    set_outer_border(sheet, 76, 92, 64, 101, medium_border, fill_white) # BP76:CW92

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

