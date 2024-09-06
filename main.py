import tkinter as tk
from tkinter import filedialog, messagebox
from lectura_datos import leer_archivo_excel, filtrar_columnas
from generar_reportes import generar_reporte_excel
from utils import imprimir_tabla, contar_casos_unicos, contar_casos_por_usuario

def ejecutar_proceso(ruta_archivo, ruta_salida):
        
        try:
            archivo_excel = leer_archivo_excel(ruta_archivo)
            tabla = filtrar_columnas(archivo_excel)
            imprimir_tabla(tabla)

            total_casos = contar_casos_unicos(archivo_excel)
            print(f"Número total de casos atendidos: {total_casos}")

            casos_bac = contar_casos_por_usuario(archivo_excel, 'topaz')
            casos_cobis = contar_casos_por_usuario(archivo_excel, 'bac')
            casos_infra = contar_casos_por_usuario(archivo_excel,'infra')
            print(f"Número de casos atendidos a Topaz: {casos_bac}")
            print(f"Número de casos atendidos a Cobis: {casos_cobis}")
            print(f"Número de casos atendidos a Infra: {casos_infra}")


            # Generar el reporte en un archivo Excel
            generar_reporte_excel(tabla, ruta_salida, total_casos, casos_bac, casos_cobis, casos_infra)

            messagebox.showinfo("Éxito", f"Reporte exportado a {ruta_salida}")
        except KeyError as e:
            messagebox.showerror("Error", f"Columnas no encontradas: {e}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def seleccionar_archivo_entrada():
        
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta_archivo:
            entrada_var.set(ruta_archivo)


def seleccionar_archivo_salida():
   
        ruta_salida = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos de Excel", "*.xlsx")])
        if ruta_salida:
            salida_var.set(ruta_salida)


def ejecutar():
    
        ruta_archivo = entrada_var.get()
        ruta_salida = salida_var.get()
        if ruta_archivo and ruta_salida:
            ejecutar_proceso(ruta_archivo, ruta_salida)
        else:
            messagebox.showwarning("Advertencia", "Debe seleccionar tanto un archivo de entrada como un archivo de salida.")

# Ejecucion de archivo de entrada y salida (Reporte)
app = tk.Tk()
app.title("Generador de Reportes Mensuales")
app.geometry("400x250")

entrada_var = tk.StringVar()
salida_var = tk.StringVar()

tk.Label(app, text="Archivo de entrada:").pack(pady=5)
tk.Entry(app, textvariable=entrada_var, width=50).pack(pady=5)
tk.Button(app, text="Seleccionar archivo", command=seleccionar_archivo_entrada).pack(pady=5)

tk.Label(app, text="Archivo de salida:").pack(pady=5)
tk.Entry(app, textvariable=salida_var, width=50).pack(pady=5)
tk.Button(app, text="Seleccionar ruta", command=seleccionar_archivo_salida).pack(pady=5)

tk.Button(app, text="Ejecutar", command=ejecutar).pack(pady=20)

app.mainloop()
