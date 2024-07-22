from tabulate import tabulate

def imprimir_tabla(df):
        """Imprimir el marco de datos en formato tabular."""
        print(tabulate(df, headers='keys', tablefmt='grid'))

def contar_casos_unicos(df):
    """Recuento total de casos únicos."""
    return df['ID'].nunique()

def contar_casos_por_usuario(df, user):
    """Recuento de casos únicos por usuario."""
    return df[df['USUARIO'].str.contains(user, case=False, na=False)]['ID'].nunique()

def tabla_pivote(df, index):
    """Generar tabla dinámica."""
    return df.pivot_table(index=index, values='ID', aggfunc='nunique')
