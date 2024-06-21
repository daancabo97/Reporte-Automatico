from tabulate import tabulate
from utils import tabla_pivote
import pandas as pd
from openpyxl.utils import get_column_letter

def ajustar_ancho_columnas(writer, sheet_name):
    """Ajustar el ancho de las columnas del archivo Excel."""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Obtén la letra de la columna
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ajustar_ancho = (max_length + 2)
        worksheet.column_dimensions[column].width = ajustar_ancho

def generar_reporte(df):
    """Generar e imprimir reportes."""
    casos_por_servicio = tabla_pivote(df, 'Servicio')
    casos_por_persona = tabla_pivote(df, 'Asignado a')
    casos_por_componente = tabla_pivote(df, 'Componente')
    casos_por_ambiente = tabla_pivote(df, 'Ambiente')

    print("Número de casos por servicio:")
    print(tabulate(casos_por_servicio, headers='keys', tablefmt='grid'))
    print("\nNúmero de casos por persona:")
    print(tabulate(casos_por_persona, headers='keys', tablefmt='grid'))
    print("\nNúmero de casos por componente:")
    print(tabulate(casos_por_componente, headers='keys', tablefmt='grid'))
    print("\nNúmero de casos por ambiente:")
    print(tabulate(casos_por_ambiente, headers='keys', tablefmt='grid'))

def generar_reporte_excel(df, ruta_salida, total_casos, casos_topaz, casos_cobis):
    """Generar reporte y exportar a un archivo Excel."""
    casos_por_servicio = tabla_pivote(df, 'Servicio')
    casos_por_persona = tabla_pivote(df, 'Asignado a')
    casos_por_componente = tabla_pivote(df, 'Componente')
    casos_por_ambiente = tabla_pivote(df, 'Ambiente')

    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Datos Filtrados', index=False)
        ajustar_ancho_columnas(writer, 'Datos Filtrados')

        resumen = pd.DataFrame({
            'Descripción': ['Total Casos Atendidos', 'Casos Topaz', 'Casos Cobis'],
            'Cantidad': [total_casos, casos_topaz, casos_cobis]
        })
        resumen.to_excel(writer, sheet_name='Resumen', index=False)
        ajustar_ancho_columnas(writer, 'Resumen')

        casos_por_servicio.to_excel(writer, sheet_name='Casos por Servicio')
        ajustar_ancho_columnas(writer, 'Casos por Servicio')

        casos_por_persona.to_excel(writer, sheet_name='Casos por Persona')
        ajustar_ancho_columnas(writer, 'Casos por Persona')

        casos_por_componente.to_excel(writer, sheet_name='Casos por Componente')
        ajustar_ancho_columnas(writer, 'Casos por Componente')

        casos_por_ambiente.to_excel(writer, sheet_name='Casos por Ambiente')
        ajustar_ancho_columnas(writer, 'Casos por Ambiente')

    print(f"Reporte exportado a: {ruta_salida}")
