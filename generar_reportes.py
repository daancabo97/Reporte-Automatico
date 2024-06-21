from tabulate import tabulate
from utils import tabla_pivote
import pandas as pd

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

    with pd.ExcelWriter(ruta_salida) as writer:
        df.to_excel(writer, sheet_name='Datos Filtrados', index=False)
        pd.DataFrame({'Total Casos Atendidos': [total_casos]}).to_excel(writer, sheet_name='Resumen', index=False)
        pd.DataFrame({'Casos Topaz': [casos_topaz], 'Casos Cobis': [casos_cobis]}).to_excel(writer, sheet_name='Resumen', startrow=2, index=False)
        casos_por_servicio.to_excel(writer, sheet_name='Casos por Servicio')
        casos_por_persona.to_excel(writer, sheet_name='Casos por Persona')
        casos_por_componente.to_excel(writer, sheet_name='Casos por Componente')
        casos_por_ambiente.to_excel(writer, sheet_name='Casos por Ambiente')

    print(f"Reporte exportado a {ruta_salida}")
