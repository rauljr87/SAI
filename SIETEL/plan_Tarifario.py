import pandas as pd
data = pd.read_excel("entrada/TarifasDedicado_Mayo_2025.xls")
data.head()
data.columns.values
resultado = (
    data.groupby(['NOMBRE COMERCIAL DEL PLAN TARIFARIO'])['CANTIDAD ABONADOS/ CLIENTES']
        .sum()
        .reset_index()
        .sort_values(by='CANTIDAD ABONADOS/ CLIENTES', ascending=False)
)
resultado
data['CANTIDAD ABONADOS/ CLIENTES'].sum()
