import os
import pandas as pd
import re

directorio = 'C:\\Python312\\excelTest2'
nombres_tarjetas = ["VISA DEBIT", "VISA", "MASTERCARD DEBIT", "MASTERCARD", "ARGENCARD", "CABAL DEBIT", "CABAL", "AMEX", "MAESTRO"]
dfs_tarjetas = {nombre: pd.DataFrame() for nombre in nombres_tarjetas}

def nombre_en_archivo(nombre_tarjeta, nombre_archivo):
    if "DEBIT" in nombre_tarjeta:
        pattern = rf'\b{re.escape(nombre_tarjeta)}(?:_|\b)'
    else:
        debit_free_pattern = rf'\b{re.escape(nombre_tarjeta)}(?:_|\b)(?!\sDEBIT)'
        pattern = debit_free_pattern if nombre_tarjeta in ["VISA", "MASTERCARD"] else rf'\b{re.escape(nombre_tarjeta)}(?:_|\b)'
    result = re.search(pattern, nombre_archivo) is not None
    print(f"Checking {nombre_archivo} for {nombre_tarjeta} using pattern {pattern}: {result}")
    return result

for nombre_tarjeta in nombres_tarjetas:
    archivos = [f for f in os.listdir(directorio) if nombre_en_archivo(nombre_tarjeta, f) and f.endswith('.xlsx')]
    print(f"Archivos encontrados para {nombre_tarjeta}: {archivos}")
    for archivo in archivos:
        ruta_completa = os.path.join(directorio, archivo)
        df_temp = pd.read_excel(ruta_completa, header=None)
        rows = df_temp.values.tolist()
        for row in rows:
            while row[0]=="":
                row=row[1:]+[""]
                  
                if not any(row):
                    break


        dfs_tarjetas[nombre_tarjeta] = pd.concat([dfs_tarjetas[nombre_tarjeta], df_temp])

for nombre_tarjeta, df in dfs_tarjetas.items():
    if not df.empty:
        output_path = f'C:\\Python312\\excelConverted\\{nombre_tarjeta}_combined.xlsx'
        df.to_excel(output_path, index=False)
        print(f'Archivo guardado: {output_path}')

