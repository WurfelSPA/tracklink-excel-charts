"""
TrackGTS Excel Charts Service

Recibe un Excel del reporte de excesos de velocidad y le agrega 2 gráficos de columnas:
- Chart 1 (Excesos de Velocidad): Columna A (Alias) + Columna B (Excesos), anclado a columna A
- Chart 2 (Velocidad Máxima): Columna A (Alias) + Columna F (Velocidad Máxima), anclado a columna F

Endpoints:
  GET  /              → health check
  POST /add-charts    → recibe Excel base64, devuelve Excel base64 con gráficos

Autenticación: header `Authorization: Bearer <API_KEY>`
La API_KEY se configura como variable de entorno en Render.
"""

import os
import io
import base64
import logging
from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

API_KEY = os.environ.get('API_KEY', '')


def check_auth():
    """Valida el header Authorization: Bearer <API_KEY>."""
    if not API_KEY:
        log.warning('API_KEY no configurada — autenticación deshabilitada')
        return True
    auth = request.headers.get('Authorization', '')
    if not auth.startswith('Bearer '):
        return False
    return auth[7:].strip() == API_KEY


def make_bar_chart(ws, title, value_col, y_axis_title, style, last_row):
    """Construye un gráfico de columnas configurado para el reporte de excesos."""
    ch = BarChart()
    ch.type = 'col'
    ch.style = style  # multicolor
    ch.title = title
    ch.x_axis.title = 'Unidad'
    ch.y_axis.title = y_axis_title
    ch.legend = None

    # Forzar visibilidad de ejes (workaround de openpyxl)
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.majorTickMark = 'out'
    ch.x_axis.minorTickMark = 'none'
    ch.y_axis.majorTickMark = 'out'
    ch.y_axis.minorTickMark = 'none'

    # Tamaño (en cm aprox según openpyxl)
    ch.width = 22
    ch.height = 13

    data = Reference(ws, min_col=value_col, min_row=5, max_row=last_row, max_col=value_col)
    cats = Reference(ws, min_col=1, min_row=6, max_row=last_row, max_col=1)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)

    return ch


@app.route('/', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'service': 'trackgts-excel-charts',
        'version': '1.0.0'
    })


@app.route('/add-charts', methods=['POST'])
def add_charts():
    if not check_auth():
        return jsonify({'error': 'Unauthorized — invalid or missing API key'}), 401

    try:
        data = request.get_json(silent=True)
        if not data or 'excelBase64' not in data:
            return jsonify({'error': 'Body debe incluir campo "excelBase64"'}), 400

        excel_bytes = base64.b64decode(data['excelBase64'])

        wb = load_workbook(io.BytesIO(excel_bytes))

        # Encontrar la hoja "Resumen 1" (puede variar nombre)
        sheet_name = None
        for name in wb.sheetnames:
            if 'resumen' in name.lower():
                sheet_name = name
                break

        if not sheet_name:
            return jsonify({
                'error': 'No se encontró hoja con "Resumen" en el nombre',
                'sheetsFound': wb.sheetnames
            }), 400

        ws = wb[sheet_name]

        # Determinar última fila con datos (después del header en fila 5)
        last_row = ws.max_row
        for row in range(last_row, 5, -1):
            if ws.cell(row=row, column=1).value:
                last_row = row
                break

        if last_row < 6:
            return jsonify({'error': 'No hay filas de datos después de la fila 5'}), 400

        rows_count = last_row - 5
        chart_anchor_row = last_row + 2  # Una fila en blanco entre datos y gráfico

        log.info(f'Procesando hoja "{sheet_name}" con {rows_count} filas de datos')

        # Gráfico 1: Excesos de Velocidad (col A + col B)
        chart1 = make_bar_chart(ws, 'Excesos de Velocidad', 2, 'Cantidad de Excesos', 10, last_row)
        ws.add_chart(chart1, f'A{chart_anchor_row}')

        # Gráfico 2: Velocidad Máxima (col A + col F)
        chart2 = make_bar_chart(ws, 'Velocidad Máxima', 6, 'Velocidad Máxima (Km/h)', 11, last_row)
        ws.add_chart(chart2, f'F{chart_anchor_row}')

        # Guardar
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        result_b64 = base64.b64encode(output.read()).decode('utf-8')

        return jsonify({
            'excelBase64': result_b64,
            'info': {
                'sheet': sheet_name,
                'rowsCount': rows_count,
                'chartAnchorRow': chart_anchor_row
            }
        })

    except Exception as e:
        log.exception('Error procesando Excel')
        return jsonify({'error': str(e), 'type': type(e).__name__}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
