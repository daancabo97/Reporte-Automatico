from tabulate import tabulate

def imprimir_tabla(df):
    """Print the dataframe in tabular format."""
    print(tabulate(df, headers='keys', tablefmt='grid'))

def contar_casos_unicos(df):
    """Count total unique cases."""
    return df['ID'].nunique()

def contar_casos_por_usuario(df, user):
    """Count unique cases by user."""
    return df[df['USUARIO'].str.contains(user, case=False, na=False)]['ID'].nunique()

def tabla_pivote(df, index):
    """Generate pivot table."""
    return df.pivot_table(index=index, values='ID', aggfunc='nunique')
