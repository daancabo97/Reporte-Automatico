import pandas as pd

def leer_archivo_excel(ruta_archivo):
    """Leer archivo excel."""
    return pd.read_excel(ruta_archivo)

def filtrar_columnas(df):
    """Filtrar las columnas requeridas del dataframe."""
    mostrar_columnas = [
        'USUARIO', 'ID', 'Fecha/Hora Asignación', 'Asignado a', 'Estado',
        'Ambiente', 'Servicio', 'Componente', 'Duracion (Hora o fraccion) Entero y decimal', 'Duracion (minutos)'
    ]
    return df[mostrar_columnas]
