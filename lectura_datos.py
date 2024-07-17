import pandas as pd

def leer_archivo_excel(ruta_archivo):
    df = pd.read_excel(ruta_archivo)
    print("Columnas disponibles en el archivo Excel:")
    print(df.columns)
    return df

def filtrar_columnas(df):
    mostrar_columnas = [
        'USUARIO', 'ID', 'Fecha/Hora Asignación', 'Fecha\nUltima Asignación', 'Fecha\nCierre', 'Dias SLA', 
        'Asignado a', 'Estado', 'Ambiente', 'Servicio', 'Componente', 
        'Duracion\n(Hora o fraccion)\nEntero y decimal', 'Duracion (minutos)'
    ]
    return df[mostrar_columnas]
