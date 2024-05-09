import os
import pandas as pd
import camelot
import aspose.pdf as ap

def convert_pdf_to_excel_with_aspose(pdf_path, excel_path):
    document = ap.Document(pdf_path)
    excel_save_options = ap.ExcelSaveOptions()
    excel_save_options.format = ap.ExcelSaveOptions.ExcelFormat.XLSX
    document.save(excel_path, excel_save_options)
    print(f'{pdf_path} convertido a Excel con Aspose en {excel_path}')
    if "MAESTRO" in pdf_path.upper():
        adjust_columns_excel(excel_path)

def adjust_columns_excel(excel_path):
    try:
        data = pd.read_excel(excel_path, header=0)
        if data.shape[1] >= 8:
            data.iloc[:, 2] =data.iloc[:,2].astype(str) + ' ' + data.iloc[:, 3].astype(str) + ' ' + data.iloc[:, 4].astype(str)
            data.iloc[:, 3] = data.iloc[:, 5]
            data.iloc[:, 4] = data.iloc[:, 6]
            data.iloc[:, 5] = data.iloc[:, 7]
            data.drop(data.columns[[6, 7]], axis=1, inplace=True)
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                data.to_excel(writer, index=False)
            print(f'Ajuste de columnas realizado en {excel_path}')
    except Exception as e:
        print(f"Error al ajustar las columnas en {excel_path}: {e}")

def convert_pdf_to_excel(pdf_path, excel_path):
    tables = camelot.read_pdf(pdf_path, flavor='lattice', pages='all')
    all_data_frames = []
    if tables.n > 0:
        for table in tables:
            df = table.df
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
            df.replace('', pd.NA, inplace=True)

            if "MAESTRO" in pdf_path.upper():
                df.iloc[:, 2] = df.iloc[:, 2].replace(r'\n', ' ', regex=True) 
            if not df.empty:
                all_data_frames.append(df)
        final_df = pd.concat(all_data_frames, ignore_index=True)
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Todas las Tablas', index=False)
        print(f'{pdf_path} convertido a {excel_path} con Camelot en una sola hoja')
    else:
        convert_pdf_to_excel_with_aspose(pdf_path, excel_path)
def process_pdfs_from_directory(pdf_directory, excel_directory):
    if not os.path.exists(excel_directory):
        os.makedirs(excel_directory)

    for filename in os.listdir(pdf_directory):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_directory, filename)
            excel_filename = filename[:-4] + '.xlsx'
            excel_path = os.path.join(excel_directory, excel_filename)
            convert_pdf_to_excel(pdf_path, excel_path)

pdf_directory = 'C:\\Python312\\pdfsPages'
excel_directory = 'C:\\Python312\\excelTest2'

process_pdfs_from_directory(pdf_directory, excel_directory)
