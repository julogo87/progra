import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
import pandas as pd
import plotly.express as px
import base64
import io

# Inicializar la aplicación Dash
app = dash.Dash(__name__)

# Layout de la aplicación
app.layout = html.Div([
    html.H1("Visualización de Programación de Vuelos QT"),
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Arrastra y suelta o ',
            html.A('Selecciona un archivo CSV')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        multiple=False
    ),
    dcc.Graph(id='flight-schedule-graph')
])

# Callback para procesar el archivo subido
@app.callback(
    Output('flight-schedule-graph', 'figure'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def update_output(contents, filename):
    if contents is None:
        return {}

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            df = pd.read_csv(io.StringIO(decoded.decode('utf-8')), sep=';')
            df['fecha_salida'] = pd.to_datetime(df['fecha_salida'])
            df['fecha_llegada'] = pd.to_datetime(df['fecha_llegada'])
            
            fig = px.timeline(df, x_start='fecha_salida', x_end='fecha_llegada', y='aeronave', color='aeronave',
                              hover_data={'numero_vuelo': True, 'origen': True, 'destino': True})
            fig.update_layout(xaxis_title='Hora', yaxis_title='Aeronave', title='Programación de Vuelos QT')
            return fig
    except Exception as e:
        print(e)
        return {}

if __name__ == '__main__':
    app.run_server(debug=True, host='0.0.0.0', port=8050)


