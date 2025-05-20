import pandas as pd
import glob
import os

# Ruta donde están los archivos Excel
ruta_archivos = r"C:\M_Workspace\python_ml_lab_n\gruposado\307 - ORIENTE\*.xlsx"

# Lista para almacenar todos los DataFrames
dataframes = []

# Leer todos los archivos Excel
archivos_excel = glob.glob(ruta_archivos)

print(f"Se encontraron {len(archivos_excel)} archivos Excel")

# Procesar cada archivo
for archivo in archivos_excel:
    print(f"Procesando: {os.path.basename(archivo)}")
    try:
        # Leer el archivo Excel
        df = pd.read_excel(archivo)
        
        # Agregar una columna con el nombre del archivo origen
        df['archivo_origen'] = os.path.basename(archivo)
        
        # Agregar el DataFrame a la lista
        dataframes.append(df)
        
    except Exception as e:
        print(f"Error procesando {archivo}: {e}")

# Unificar todos los DataFrames
if dataframes:
    df_unificado = pd.concat(dataframes, ignore_index=True)
    
    # Guardar el archivo unificado
    nombre_archivo_salida = "excel_unificado.xlsx"
    df_unificado.to_excel(nombre_archivo_salida, index=False)
    
    print(f"\n✓ Archivo unificado creado: {nombre_archivo_salida}")
    print(f"Total de filas: {len(df_unificado)}")
    print(f"Total de columnas: {len(df_unificado.columns)}")
else:
    print("No se pudieron procesar archivos Excel")