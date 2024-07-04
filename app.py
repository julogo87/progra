import pandas as pd
import matplotlib.pyplot as plt

# Datos de ejemplo
data = {
    'fecha_salida': pd.date_range(start='2023-07-01', periods=10, freq='H'),
    'fecha_llegada': pd.date_range(start='2023-07-01', periods=10, freq='H') + pd.Timedelta(hours=1),
    'aeronave': ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT', 'N330QT', 'N331QT', 'N332QT'],
    'numero_vuelo': [f'QT{i}' for i in range(10)],
    'origen': ['A'] * 10,
    'destino': ['B'] * 10
}

df = pd.DataFrame(data)

# Convertir las columnas de fecha a tipo datetime
df['fecha_salida'] = pd.to_datetime(df['fecha_salida'])
df['fecha_llegada'] = pd.to_datetime(df['fecha_llegada'])

# Crear la figura y el eje
fig, ax = plt.subplots(figsize=(20, 10))

# Orden de aeronaves
order = ['N330QT', 'N331QT', 'N332QT', 'N334QT', 'N335QT', 'N336QT', 'N337QT']
df['aeronave'] = pd.Categorical(df['aeronave'], categories=order, ordered=True)
df = df.sort_values('aeronave', ascending=False)

# Iterar sobre las aeronaves y agregar los vuelos al gráfico
for i, aeronave in enumerate(reversed(order)):
    vuelos_aeronave = df[df['aeronave'] == aeronave]
    for _, vuelo in vuelos_aeronave.iterrows():
        start = vuelo['fecha_salida']
        duration = vuelo['fecha_llegada'] - vuelo['fecha_salida']
        rect_height = 0.2  # Ajustar el ancho de las franjas
        ax.broken_barh([(start, duration)], (i - rect_height/2, rect_height), facecolors='red')  # Cambio de color a rojo

# Formato del eje Y
ax.set_yticks(range(len(order)))
ax.set_yticklabels(reversed(order))
ax.set_ylim(-1, len(order))

# Formato del eje X
plt.xticks(rotation=45, fontsize=10)

# Etiquetas y título
plt.xlabel('Hora')
plt.ylabel('Aeronave')
plt.title('Programación de Vuelos QT')

# Mostrar el gráfico
plt.show()
