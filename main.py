import argparse
from lectura_datos import leer_archivo_excel, filtrar_columnas
from generar_reportes import generar_reporte, generar_reporte_excel
from utils import imprimir_tabla, contar_casos_unicos, contar_casos_por_usuario

def parseo_argumentos():
    """Parseo de argumentos del archivo excel."""
    parser = argparse.ArgumentParser(description='Procesando archivo')
    parser.add_argument('ruta_archivo', type=str, help='La ruta del archivo de entrada')
    parser.add_argument('ruta_salida', type=str, help='La ruta del archivo de salida')
    return parser.parse_args()

def main():
    args = parseo_argumentos()
    archivo_excel = leer_archivo_excel(args.ruta_archivo)

    tabla = filtrar_columnas(archivo_excel)
    imprimir_tabla(tabla)

    total_casos = contar_casos_unicos(archivo_excel)
    print(f"Número total de casos atendidos: {total_casos}")

    casos_topaz = contar_casos_por_usuario(archivo_excel, 'topaz')
    casos_cobis = contar_casos_por_usuario(archivo_excel, 'bac')
    print(f"Número de casos atendidos a Topaz: {casos_topaz}")
    print(f"Número de casos atendidos a Cobis: {casos_cobis}")

    generar_reporte(archivo_excel)

    # Generar el reporte en un archivo Excel
    generar_reporte_excel(tabla, args.ruta_salida, total_casos, casos_topaz, casos_cobis)

if __name__ == "__main__":
    main()
