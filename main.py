import os
import shutil
from datetime import datetime
import openpyxl
import subprocess
import platform
import argparse

def main():
    # Parsear argumentos de línea de comandos
    parser = argparse.ArgumentParser(description='Generador de Actas desde Excel')
    parser.add_argument('--fila-inicio', type=int, default=2,
                        help='Fila de inicio (default: 2)')
    parser.add_argument('--fila-fin', type=int, default=None,
                        help='Fila de fin (default: última fila)')
    
    args = parser.parse_args()
    
    resumen_path = 'input/Resumen.xlsx'
    fila_inicio = args.fila_inicio
    fila_fin = args.fila_fin
    
    # Obtener la fecha de hoy
    today = datetime.now()
    date_str = today.strftime('%Y%m%d')  # Formato YYYYMMDD

    # Crear la carpeta en resultados con la fecha
    result_dir = os.path.join('resultados', date_str)
    os.makedirs(result_dir, exist_ok=True)

    # Leer el archivo Resumen.xlsx
    wb_resumen = openpyxl.load_workbook(resumen_path)
    sheet_resumen = wb_resumen.active
    
    # Si no se especifica fila_fin, usar la última fila con datos
    if fila_fin is None:
        fila_fin = sheet_resumen.max_row

    # Procesar filas en el rango especificado
    for row_idx, row in enumerate(sheet_resumen.iter_rows(min_row=fila_inicio, max_row=fila_fin, values_only=True), start=fila_inicio):
        # Asumir columnas: 0=ID, 1=Nombres, 2=Apellidos, 3=Pais, 4=edad, 5=Servicios
        id_val = row[0]
        nombres = row[1]
        apellidos = row[2]
        pais = row[3]
        edad = row[4]
        servicios = row[5] if row[5] else ""  # Servicios como string

        # Generar timestamp único para cada archivo (segundos y milisegundos)
        timestamp = date_str + "_" + datetime.now().strftime('%H%M%S%f')  # HHMMSSffffff

        # Nombre del archivo: <ID>_Acta_<timestamp>.xlsx
        acta_name = f"{id_val}_Acta_{timestamp}.xlsx"
        acta_path = os.path.join(result_dir, acta_name)

        # Copiar el modelo
        modelo_path = 'modelo/ModeloActa.xlsx'
        shutil.copy(modelo_path, acta_path)

        # Editar el archivo copiado
        wb_acta = openpyxl.load_workbook(acta_path)
        sheet_acta = wb_acta.active

        # Rellenar campos
        sheet_acta['C5'] = nombres
        sheet_acta['C6'] = apellidos
        sheet_acta['C7'] = edad
        sheet_acta['C8'] = pais

        # Servicios: colocar X en las celdas correspondientes
        if 'Gasfiteria' in servicios:
            sheet_acta['B12'] = 'X'
        if 'Plomeria' in servicios:
            sheet_acta['B14'] = 'X'
        if 'Electricista' in servicios:
            sheet_acta['B16'] = 'X'

        # Guardar el archivo Excel
        wb_acta.save(acta_path)

        # Convertir a PDF usando LibreOffice
        pdf_name = acta_name.replace('.xlsx', '.pdf')
        pdf_path = os.path.join(result_dir, pdf_name)
        
        try:
            # Detectar ubicación de LibreOffice según el SO
            if platform.system() == "Windows":
                libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            elif platform.system() == "Darwin":  # macOS
                libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            else:  # Linux
                libreoffice_path = "soffice"
            
            # Convertir Excel a PDF
            subprocess.run(
                [libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', result_dir, acta_path],
                check=True,
                capture_output=True
            )
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            print(f"Advertencia: No se pudo convertir {acta_name} a PDF. Asegúrate de que LibreOffice esté instalado.")

    print("Proceso completado.")

if __name__ == "__main__":
    main()