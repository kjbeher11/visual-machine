import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import pandas as pd

# Cargar los datos desde el archivo Excel
df = pd.read_excel('Indices.xlsx')

departamentos_co = {
    "BOGOTA D.C.": {"latitud": 4.6097, "longitud": -74.0817},
    "ANTIOQUIA": {"latitud": 6.4889, "longitud": -75.5700},
    "VALLE": {"latitud": 3.9000, "longitud": -76.9667},
    "SANTANDER": {"latitud": 7.5000, "longitud": -73.0000},
    "CUNDINAMARCA": {"latitud": 4.5710, "longitud": -74.1322},
    "ATLANTICO": {"latitud": 10.4747, "longitud": -74.9302},
    "RISARALDA": {"latitud": 4.0933, "longitud": -75.8480},
    "CALDAS": {"latitud": 5.5000, "longitud": -75.6667},
    "NARINO": {"latitud": 1.2000, "longitud": -77.0000},
    "SUCRE": {"latitud": 9.2915, "longitud": -75.1902},
    "TOLIMA": {"latitud": 4.3333, "longitud": -75.7500},
    "BOYACA": {"latitud": 5.6633, "longitud": -72.4810},
    "BOLIVAR": {"latitud": 10.2500, "longitud": -75.5000},
    "MAGDALENA": {"latitud": 10.5042, "longitud": -74.2274},
    "CAUCA": {"latitud": 2.7050, "longitud": -76.8260},
    "SAN ANDRES Y PROVIDENCIA": {"latitud": 12.5833, "longitud": -81.7000},
    "CORDOBA": {"latitud": 8.4324, "longitud": -75.8894},
    "LA GUAJIRA": {"latitud": 11.3548, "longitud": -72.5205},
    "CAQUETA": {"latitud": 1.7479, "longitud": -75.6102},
    "META": {"latitud": 3.8833, "longitud": -73.6333},
    "HUILA": {"latitud": 2.7045, "longitud": -75.9613},
    "NORTE DE SANTANDER": {"latitud": 7.8833, "longitud": -72.5000},
    "CESAR": {"latitud": 10.0739, "longitud": -73.6993},
    "CASANARE": {"latitud": 5.8956, "longitud": -71.7465},
    "QUINDIO": {"latitud": 4.5352, "longitud": -75.6095},
    "PUTUMAYO": {"latitud": 1.8000, "longitud": -76.5000},
    "AMAZONAS": {"latitud": -1.4429, "longitud": -71.5724},
    "ARAUCA": {"latitud": 6.5489, "longitud": -71.1730},
    "CHOCO": {"latitud": 5.1500, "longitud": -76.6500},
    "GUAVIARE": {"latitud": 3.1316, "longitud": -70.1783},
    "GUAINIA": {"latitud": 3.3600, "longitud": -67.2100},
}

# Crear la aplicación Dash
app = dash.Dash(__name__)
server = app.server

# Crear el layout de la aplicación
app.layout = html.Div([
    html.Div([
        html.H1("Visualización Financiera de PYMES en Colombia", style={'textAlign': 'center', 'color': '#4a4a4a'}),
        dcc.Slider(
            id='year-slider',
            min=df['Fecha de Corte'].min(),
            max=df['Fecha de Corte'].max(),
            value=df['Fecha de Corte'].min(),
            marks={str(year): str(year) for year in df['Fecha de Corte'].unique()},
            step=None,
            tooltip={"placement": "bottom", "always_visible": True},
            included=False,
            updatemode='drag'
        ),
    ], style={'margin': '20px', 'backgroundColor': '#f9f9f9', 'borderRadius': '8px', 'padding': '20px'}),

    # Tabla interactiva
    dash_table.DataTable(
        id='table',
        columns=[{"name": i, "id": i} for i in df.columns],
        data=df.to_dict('records'),  # Mostrar todos los registros inicialmente
        style_table={'overflowX': 'auto', 'border': 'thin lightgrey solid'},
        style_cell={'textAlign': 'center', 'padding': '10px', 'fontSize': '14px'},
        style_header={'backgroundColor': '#eaeaea', 'fontWeight': 'bold', 'color': '#333'},
        style_data={'backgroundColor': 'white', 'color': '#666'},
        page_size=5,  # Número de filas por página
        fixed_rows={'headers': True},
    ),

    # Primera fila con gráficos de barras y pastel
    html.Div([
        # Diagrama de barras
        dcc.Graph(id='bar-plot', style={'display': 'inline-block', 'width': '48%'}),

        # Gráfico de pastel
        dcc.Graph(id='pie-chart', style={'display': 'inline-block', 'width': '48%'}),
    ], style={'display': 'flex', 'justify-content': 'space-between', 'margin': '20px 0'}),

    # Segunda fila con boxplot múltiple y mapa de calor
    html.Div([
        # Boxplot múltiple (ROE vs ROA)
        dcc.Graph(id='boxplot', style={'display': 'inline-block', 'width': '48%'}),

        # Mapa de calor
        dcc.Graph(id='heatmap', style={'display': 'inline-block', 'width': '48%'})
    ], style={'display': 'flex', 'justify-content': 'space-between', 'margin': '20px 0'})
], style={'backgroundColor': '#f4f4f4', 'padding': '20px'})

# Callback para actualizar los gráficos
@app.callback(
    Output('bar-plot', 'figure'),
    Output('pie-chart', 'figure'),
    Output('boxplot', 'figure'),
    Output('heatmap', 'figure'),
    Input('year-slider', 'value')
)
def update_graphs(selected_year):
    # Filtrar el DataFrame por el año seleccionado
    filtered_df = df[df['Fecha de Corte'] == selected_year]

    # Contar el número de registros por departamento
    department_counts = filtered_df['Departamento de la dirección del domicilio'].value_counts().reset_index()
    department_counts.columns = ['Departamento', 'Conteo']
    department_counts = department_counts.sort_values(by='Conteo', ascending=False)

    # Crear el diagrama de barras
    bar_fig = px.bar(department_counts, x='Departamento', y='Conteo',
                     title=f"Conteo de Registros por Departamento en el Año {selected_year}",
                     labels={'Departamento': 'Departamento', 'Conteo': 'Cantidad de Registros'},
                     height=500, template='plotly_white')

    # Contar el número de registros por tipo societario
    type_counts = filtered_df['Tipo societario'].value_counts().reset_index()
    type_counts.columns = ['Tipo Societario', 'Conteo']

    # Crear el gráfico de pastel
    pie_fig = px.pie(type_counts, names='Tipo Societario', values='Conteo',
                     title=f"Distribución de PYMES por Tipo Societario en el Año {selected_year}",
                     template='plotly_white')

    # Crear el boxplot múltiple (ROE vs ROA)
    boxplot_df = filtered_df[['ROE', 'ROA']]
    boxplot_df = pd.melt(boxplot_df, var_name='Tipo', value_name='Monto')
    boxplot_fig = px.box(boxplot_df, x='Tipo', y='Monto',
                         title="Distribución de ROE y ROA",
                         labels={'Tipo': 'Categoría', 'Monto': 'Monto'},
                         template='plotly_white')

    # Agregar coordenadas geográficas a los departamentos
    department_counts['Latitud'] = department_counts['Departamento'].map(lambda dept: departamentos_co[dept]['latitud'])
    department_counts['Longitud'] = department_counts['Departamento'].map(lambda dept: departamentos_co[dept]['longitud'])

    # Crear el mapa de calor usando las coordenadas
    heatmap_fig = px.scatter_geo(department_counts, lat='Latitud', lon='Longitud', size='Conteo',
                                 title="Mapa de Calor de Empresas por Departamento",
                                 hover_name='Departamento', size_max=50,
                                 projection="natural earth", scope='south america',
                                 fitbounds="locations", template='plotly_white')  # Zoom ajustado a Colombia

    return bar_fig, pie_fig, boxplot_fig, heatmap_fig

# Ejecutar la aplicación
if __name__ == '__main__':
    app.run_server()