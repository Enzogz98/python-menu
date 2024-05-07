import pandas as pd
import os

# Headers predeterminados, cada archivo puede tener menos o más dependiendo del caso.
fixed_headers = [
    "Trx", "Fecha Pres Fecha", "Term/Lote/Cupon", "Tarj", "Plan Cuota", "T F", 
    "T.N.A. %", "Ventas con/Dto.", "Ventas sin/Dto.", "Dto. Arancel", 
    "Dto. Financ.", "Cod. Rechazo Mot. contrap."
]

def convert_to_float(x):
    if pd.notna(x) and isinstance(x, str):
        try:
            return float(x.replace('.', '').replace(',', '.'))
        except ValueError:
            return x  # Retorna el valor original si la conversión falla
    return x

def separate_lote_cupon(data):
    split_content = data.iloc[:, 2].str.split(expand=True)
    data.insert(3, 'lote', split_content[1])
    data.insert(4, 'cupon', split_content[2]) if split_content.shape[1] > 1 else None
    data.iloc[:, 2] = split_content[0] if split_content.shape[1] > 2 else None
    return data

def clean_empty_cells(data, start_col, sheet_name):
    for col in range(start_col, len(data.columns)):
        data.iloc[:, col] = data.iloc[:, col].apply(lambda x: x if pd.notna(x) and str(x).strip() != '' else None)
    if 'MAESTRO' in sheet_name:
        data['Term/Lote/Cupon'] = data['Term/Lote/Cupon'].apply(lambda x: x if pd.notna(x) and str(x).strip() != '' else "Default Value")
    return data

def apply_numeric_conversion(data):
    for col in data.columns:
        data[col] = data[col].apply(convert_to_float)
    return data

def filter_excel(file_path):
    try:
        data = pd.read_excel(file_path, header=1)
        
        # Eliminar la 13ª columna si existe
        if data.shape[1] > 12:
            data = data.iloc[:, :-1]
        
        # Ajusta los nombres de las columnas basados en la cantidad de columnas presentes en el archivo
        current_headers = fixed_headers[:min(len(fixed_headers), len(data.columns))]
        data.columns = current_headers
        
        if 'MAESTRO' in file_path:
            data = separate_lote_cupon(data)
            data = clean_empty_cells(data, 2, 'MAESTRO')
        else:
            data = clean_empty_cells(data, 3, 'Other')
        
        # Aplica la conversión numérica a todas las celdas
        data = apply_numeric_conversion(data)
        
        if data.iloc[:, 0].isin(['Plan cuota', 'Venta ctdo']).any():
            filtered_data = data[data.iloc[:, 0].isin(['Plan cuota', 'Venta ctdo'])]
            return filtered_data
        else:
            print(f"No se encontraron entradas válidas en la primera columna del archivo: {file_path}")
            return None
    except Exception as e:
        print(f"Error procesando el archivo {file_path}: {e}")
        return None

def process_folder(folder_path):
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    output_folder_path = os.path.join(folder_path, 'excelCrudo')
    os.makedirs(output_folder_path, exist_ok=True)
    
    for file in files:
        file_path = os.path.join(folder_path, file)
        filtered_data = filter_excel(file_path)
        
        if filtered_data is not None:
            output_path = os.path.join(output_folder_path, f'filtered_{file}')
            filtered_data.to_excel(output_path, index=False)
            print(f'Archivo filtrado guardado en: {output_path}')

folder_path = r'C:\Python312\excelConverted'
process_folder(folder_path)
