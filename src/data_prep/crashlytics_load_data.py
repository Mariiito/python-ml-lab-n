from pathlib import Path
import pandas as pd
import glob
import os

def merge_csv_prefix():
    """
        _summary_
    """
    
    # Directorio de destino
    output_dir = Path('/data/processed')
    """
        En caso de no existir, se crea:
        output_dir.mkdir(parents=True, exist_ok=True)
    """
    
    prefixes = ['acc', 'pers', 'veh']
    
    for prefix in prefixes:
        print("")
        
        # Se buscan los archivos CSV que comienzan con el prefijo
        pattern = f"{prefix}*.csv"
        csv_files = glob.glob(pattern)
        
        if not csv_files:
            print("No existen archivos CSV con el prefijo: {prefix}")
            continue
        
        print(f"CSV encontrados con el prefijo {prefix}:")
        
        # Se almacenan los dataframes en una lista
        dataframes = []
        
        for file in csv_files:
            try:
                df = pd.read_csv(file)
            except Exception as e:
                print(f"Error al leer {file}_ {e}")
                
        if dataframes:
            # Unión de los dataframes
            merged_df = pd.concat(dataframes, ignore_index=True)
            
            # Guardar dataset combinado
            output_file = output_dir / f"{prefix}_merged.csv"
            merged_df.to_csv(output_file, index=False)
        else:
            print(f"No se pudo procesar nigún archivo con el prefijo {prefix}")
            
def main():
    merge_csv_prefix()
    print("\n=== Proceso finalizado ===")
    
    
if __name__ == "__main__":
    main()
                