import os
import pandas as pd

# Obtener el nombre de usuario del PC
user_directory = os.path.expanduser("~")

# Lista de archivos que deseas leer
archivos = [
    "ZBD Contenedores.xlsx",
    "ZBD ContR3.XLSX",
    "ZBD DTs.xlsx",
    "ZBD DTsR3.xlsx",
    "ZBD IMO.xlsx",
    "Sector.XLSX"  # Agrega el nuevo archivo aquí
]

# Crear una lista de rutas completas de los archivos
rutas_archivos = [os.path.join(user_directory, "OneDrive - Inchcape", "Macro Memo", archivo) for archivo in archivos]

# Leer los archivos con pandas
dataframes = [pd.read_excel(archivo) for archivo in rutas_archivos]

# Ahora, dataframes contiene los datos de los archivos Excel, incluyendo "Sector.XLSX"
df_contenedores = dataframes[0]
df_sector = dataframes[5]
dfcontR3 = dataframes[1]
df_dts = dataframes[2]
df_dtsr3 = dataframes[3]
dfimo = dataframes[4]

columnas_a_mantener3 = [
    'Moneda Orden Compra',
    'Entrega entrante',
    'Nro DT',
    'Ref. Prefijo embarque',
    'Código marca/producto',
    'Vía (Texto)',
    'Nombre de proveedor',
    'Proveedor',
    'Nave/Aerolínea',
    'Documento de embarque',
    'FE.ATA',
    'Contenedor',
    'Material OC',
    'Material proveedor',
    'Descripción material',
    'Incoterm'
]

# Seleccionar las columnas específicas en df_contenedores
dfcontR3 = dfcontR3[columnas_a_mantener3]

columnas_a_mantener2 = [
    'Moneda Orden Compra',
    'Entrega entrante',
    'Nro DT',
    'Ref. Prefijo embarque',
    'Código marca/producto',
    'Vía (Texto)',
    'Nombre de proveedor',
    'Proveedor',
    'Nave/Aerolínea',
    'Documento de embarque',
    'FE.ATA',
    'Contenedor',
    'Material OC',
    'Material proveedor',
    'Descripción material',
    'Incoterm'
]

# Seleccionar las columnas específicas en df_contenedores
df_contenedores = df_contenedores[columnas_a_mantener2]
columnas_a_mantener1 = [
    'Nro. DT',
    'Cant. Factura',
    'Documento de embarque',
    'Fe. ATA',
    'Marca/Producto',
    'Moneda',
    'Nave / Aerolínea',
    'Nombre Proveedor',
    'País Origen',
    'Proveedor',
    'Ref. Prefijo Emb.',
    'Referencia',
    'Valor Fact.',
    'Vía (Texto)'
]

# Seleccionar las columnas específicas en df_dtsr3
df_dtsr3 = df_dtsr3[columnas_a_mantener1]

columnas_a_mantener4 = [
    'Nro. DT',
    'Cant. Factura',
    'Documento de embarque',
    'Fe. ATA',
    'Marca/Producto',
    'Moneda',
    'Nave / Aerolínea',
    'Nombre Proveedor',
    'País Origen',
    'Proveedor',
    'Ref. Prefijo Emb.',
    'Referencia',
    'Valor Fact.',
    'Vía (Texto)'
]
# Seleccionar columnas específicas y asegurar tipos de datos consistentes para unión y concatenación
df_dts = df_dts[columnas_a_mantener4]
dfcontR3['Nro DT'], df_dtsr3['Nro. DT'], df_contenedores['Nro DT'], df_dts['Nro. DT'] = dfcontR3['Nro DT'].astype(str), df_dtsr3['Nro. DT'].astype(str), df_contenedores['Nro DT'].astype(str), df_dts['Nro. DT'].astype(str)

# Concatenar DataFrames de DTs y contenedores, respectivamente
dt = pd.concat([df_dts, df_dtsr3], ignore_index=True)
cont = pd.concat([df_contenedores, dfcontR3], ignore_index=True)
cont['Material OC'] = cont['Material OC'].astype(str)
# Convertir columnas de 'Proveedor' a string para futuros merges
dt['Proveedor'], cont['Proveedor'], df_sector['CÓDIGO PROVEEDOR'] = dt['Proveedor'].astype(str), cont['Proveedor'].astype(str), df_sector['CÓDIGO PROVEEDOR'].astype(str)
dfimo['MATERIAL S4 2'] = dfimo['MATERIAL S4 2'].astype(str)
# cont['Nro DT'] = cont['Nro DT'].astype(str)
# dt['Nro. DT'] = dt['Nro. DT'].astype(str)

merged_df = cont.merge(dt[['Nro. DT', 'Referencia', 'Valor Fact.']], left_on='Nro DT', right_on='Nro. DT', how='left')
merged_df['Nro DT'] = merged_df['Nro DT'].astype(str)

# renault_df = dt[dt['Nro. DT'] == '22649']
df1 = df_dtsr3[df_dtsr3['Nro. DT'] == '22649']
a = merged_df[merged_df['Nro DT'] == '22649']
# Realizar el merge
merged_df = pd.merge(merged_df, dfimo[['MATERIAL S4 2', 'REQUIERE CDA', 'MOTIVO CDA', 'NÚMERO UN']], 
                           left_on='Material OC', right_on='MATERIAL S4 2', 
                           how='left')

# Opcional: Si después del merge no necesitas mantener la columna 'MATERIAL S4 2' en el DataFrame resultante, puedes eliminarla
merged_df.drop(columns=['MATERIAL S4 2'], inplace=True)
# Realizar el merge
merged_df = pd.merge(merged_df, df_sector[['CÓDIGO PROVEEDOR', 'TIPO PROVEEEDOR', 'CBE', 'INCOTERMS', 'Almacén / Bodega']], 
                           left_on='Proveedor', right_on='CÓDIGO PROVEEDOR', 
                           how='left')

# Opcional: Si después del merge no necesitas mantener la columna 'CÓDIGO PROVEEDOR' en el DataFrame resultante, puedes eliminarla
merged_df.drop(columns=['CÓDIGO PROVEEDOR'], inplace=True)
columnas_con_nan = ['REQUIERE CDA', 'MOTIVO CDA', 'NÚMERO UN', 'Referencia', 'Valor Fact.','TIPO PROVEEEDOR', 'CBE', 'INCOTERMS', 'Almacén / Bodega']

# Reemplazar NaN con "Sin info" en las columnas mencionadas
merged_df[columnas_con_nan] = merged_df[columnas_con_nan].fillna("Sin info")
df_app = merged_df

import pandas as pd
# Supongamos que tienes un DataFrame llamado df_app

# Lista de nombres de columnas actual
nombres_columnas = df_app.columns.tolist()

# Reorganizar el orden de las columnas
nombres_columnas.insert(0, nombres_columnas.pop(nombres_columnas.index('Nro DT')))  # Mueve 'Nro DT' a la primera posición
nombres_columnas.insert(3, nombres_columnas.pop(nombres_columnas.index('Nombre de proveedor')))  # Mueve 'Proveedor' a la cuarta posición

# Reorganizar las columnas en el DataFrame
df_app = df_app[nombres_columnas]

# Lista de columnas que deseas convertir a tipo texto (str)
columnas_a_texto = ['Nro DT', 'Moneda Orden Compra', 'Entrega entrante', 'Nombre de proveedor',
                    'Código marca/producto', 'Vía (Texto)', 'Proveedor', 'Nave/Aerolínea',
                    'Documento de embarque', 'Contenedor', 'Material OC',
                    'Material proveedor', 'Descripción material', 'Incoterm', 'TIPO PROVEEEDOR', 'CBE', 'INCOTERMS', 'Almacén / Bodega',
                    'REQUIERE CDA', 'MOTIVO CDA', 'NÚMERO UN', 'Referencia', 'Valor Fact.']

# Convertir las columnas a tipo texto (str)
df_app[columnas_a_texto] = df_app[columnas_a_texto].astype(str)
columnas_a_mantener62 = [
    'Nro. DT',
    'Fe. ATA'
]
df_concatenado = pd.concat([df_dts[columnas_a_mantener62], df_dtsr3[columnas_a_mantener62]])
# Seleccionar solo las columnas 'Nro. DT' y 'Fe. ATA'
df_ata = df_concatenado[['Nro. DT', 'Fe. ATA']].drop_duplicates()
df_ata['Nro. DT'] = df_ata['Nro. DT'].astype(str)
df_ata.rename(columns={'Fe. ATA': 'Fe.ATA'}, inplace=True)
df_ata.rename(columns={'Nro. DT': 'Nro DT'}, inplace=True)
# Hacer el merge basado en la columna 'Nro DT'
merged_df2 = pd.merge(df_app, df_ata, on='Nro DT', how='left')
# Filtrar las filas donde 'FE.ATA' en df_app es NaN
filas_nan = merged_df2['FE.ATA'].isna() | (merged_df2['FE.ATA'] == '')
# Rellenar los valores NaN en 'FE.ATA' de df_app con los valores correspondientes de df_ata
merged_df2.loc[filas_nan, 'FE.ATA'] = merged_df2.loc[filas_nan, 'Fe.ATA']

# Calcular los índices de las columnas a eliminar (1 y 10 desde el final)
index_1_from_end = -1  # La última columna (la primera desde el final)
index_10_from_end = -11  # La décima columna desde el final

# Obtener los nombres de las columnas a eliminar
column_name_1_from_end = merged_df2.columns[index_1_from_end]
column_name_10_from_end = merged_df2.columns[index_10_from_end]

# Eliminar las columnas del DataFrame
merged_df2 = merged_df2.drop(columns=[column_name_1_from_end, column_name_10_from_end])

# # Eliminar la columna 'Fe.ATA' que se unió de df_ata
# merged_df2.drop(columns=['Fe.ATA'], inplace=True)

# Ahora merged_df2 contiene los datos combinados con las FE.ATA llenadas donde sea posible
merged_df2.drop_duplicates
df_app = merged_df2
# Lista actualizada de columnas que deseas convertir a tipo texto (str), asegurándonos de que todas las columnas mencionadas estén incluidas.
columnas_a_texto = [
    'Nro DT', 'Moneda Orden Compra', 'Entrega entrante', 'Nombre de proveedor',
    'Ref. Prefijo embarque', 'Código marca/producto', 'Vía (Texto)', 'Proveedor', 
    'Nave/Aerolínea', 'Documento de embarque', 'Contenedor', 'Material OC',
    'Material proveedor', 'Descripción material', 'Incoterm', 'Referencia', 
    'Valor Fact.', 'REQUIERE CDA', 'MOTIVO CDA', 'NÚMERO UN', 'TIPO PROVEEEDOR', 
    'CBE', 'INCOTERMS', 'Almacén / Bodega'
]

# Asegurarte de que df_app es el DataFrame correcto al que deseas aplicar estos cambios.
# Convertir las columnas especificadas a tipo texto (str)
df_app[columnas_a_texto] = df_app[columnas_a_texto].astype(str)

df_app.drop_duplicates
df_app.rename(columns={'Moneda Orden Compra': 'MONEDA'}, inplace=True)
import pandas as pd
# Obtener la carpeta de inicio del usuario actual
user_folder = os.path.expanduser("~")

# Crear la ruta completa
# ruta = os.path.join(user_folder, "Inchcape", "Planificación y Compras Chile - Documentos", "Planificación y Compras Aftermarket", "Transitos", "COMMODITIES", "PA DT 31-03", "Macro Memo", "df_app.xlsx")
ruta = os.path.join(user_folder, "OneDrive - Inchcape", "Macro Memo", "df_app.xlsx")


# Guardar el DataFrame en formato Excel en la ruta especificada
df_app.to_excel(ruta, index=False)
import pandas as pd
import os
df_app.drop_duplicates
# Asumiendo que df_app es tu DataFrame y ya has realizado las operaciones mencionadas sobre él.

# Asumiendo que df_app es tu DataFrame y ya has realizado las operaciones mencionadas sobre él.

# Especifica la ruta donde quieres guardar tu archivo CSV. Ajusta el nombre del archivo según necesites.
user_folder = os.path.expanduser("~")
ruta_csv = os.path.join(user_folder, "OneDrive - Inchcape", "Macro Memo", "df_app.csv")

# Guardar el DataFrame en un archivo CSV.
# Asegúrate de ajustar los parámetros como sep, index, y encoding si es necesario para tu aplicación.
df_app.to_csv(ruta_csv, sep=';', index=False, encoding='utf-8-sig')
