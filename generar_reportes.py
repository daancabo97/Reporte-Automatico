from tabulate import tabulate
from utils import tabla_pivote
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.chart import BarChart, Reference

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Ajustar el formato de las columnas
def ajustar_ancho_columnas(writer, sheet_name):
        
        workbook = writer.book
        worksheet = workbook[sheet_name]
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  
            for celda in col:
                try:
                    if len(str(celda.value)) > max_length:
                        max_length = len(str(celda.value))
                except:
                    pass
            ajustar_anchura = (max_length + 2)
            worksheet.column_dimensions[column].width = ajustar_anchura

# Generar gráficas de barras
def generar_grafica_barras(worksheet, data_range, categoria_range, titulo, celda):
    
            chart = BarChart()
            chart.title = titulo
            chart.style = 10
            chart.y_axis.title = 'Casos'

            data = Reference(worksheet, min_col=data_range['min_col'], min_row=data_range['min_row'],
                            max_col=data_range['max_col'], max_row=data_range['max_row'])
            
            categories = Reference(worksheet, min_col=categoria_range['min_col'], min_row=categoria_range['min_row'],
                                max_col=categoria_range['max_col'], max_row=categoria_range['max_row'])

            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            chart.shape = 4
            chart.width = 30  # Ajustar el ancho de la gráfica
            chart.height = 15  # Ajustar la altura de la gráfica
            worksheet.add_chart(chart, celda)

# Imprimir reporte en consola
def generar_reporte(df):
    
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

# Generar reporte en Excel
def generar_reporte_excel(df, ruta_salida, total_casos, casos_topaz, casos_cobis):
    
            casos_por_servicio = tabla_pivote(df, 'Servicio')
            casos_por_persona = tabla_pivote(df, 'Asignado a')
            casos_por_componente = tabla_pivote(df, 'Componente')
            casos_por_ambiente = tabla_pivote(df, 'Ambiente')
            
            # Obtener los tres casos que más tiempo han tomado en sacar
            df['Duracion (minutos)'] = pd.to_numeric(df['Duracion (minutos)'], errors='coerce')
            casos_mas_demorados = df.nlargest(170, 'Duracion (minutos)')

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

                # Añadir gráfica de barras
                workbook = writer.book
                
                hojas_con_graficas = {
                    'Casos por Servicio': casos_por_servicio,
                    'Casos por Persona': casos_por_persona,
                    'Casos por Componente': casos_por_componente,
                    'Casos por Ambiente': casos_por_ambiente,
                    'Resumen': resumen
                }

                for nombre_hoja, data in hojas_con_graficas.items():
                    worksheet = workbook[nombre_hoja]
                    data_range = {
                        'min_col': 2,
                        'min_row': 1,
                        'max_col': 2,
                        'max_row': len(data) + 1
                    }
                    categoria_range = {
                        'min_col': 1,
                        'min_row': 2,
                        'max_col': 1,
                        'max_row': len(data) + 1
                    }
                    generar_grafica_barras(worksheet, data_range, categoria_range, f'{nombre_hoja}', 'E5')

                # Añadir los tres casos más Demorados
                casos_mas_demorados.to_excel(writer, sheet_name='Casos Más Demorados', index=False)
                ajustar_ancho_columnas(writer, 'Casos Más Demorados')

                # Añadir gráfica de barras para Casos Más Demorados
                worksheet = writer.sheets['Casos Más Demorados']
                data_range = {
                    'min_col': 13,  # Columna de 'Duracion (minutos)'
                    'min_row': 1,
                    'max_col': 13,
                    'max_row': len(casos_mas_demorados) + 1
                }
                categoria_range = {
                    'min_col': 2,  # Columna de 'ID'
                    'min_row': 2,
                    'max_col': 2,
                    'max_row': len(casos_mas_demorados) + 1
                }
                generar_grafica_barras(worksheet, data_range, categoria_range, 'Casos Más Demorados', 'E5')

            print(f"Se ha exportado el reporte en la ruta: {ruta_salida}")
