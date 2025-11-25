import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import os

# --- CONFIGURACIÓN ---
INPUT_DIR = "input"
OUTPUT_DIR = "output"
NOMBRE_ARCHIVO_ORIGEN = "archivo.xlsx"
ARCHIVO_SALIDA = "datos_transformados.xlsx"
HOJA_DATOS = "RECLAM. FB"
HOJA_FECHA = "RESULTADOS FB"

# Definir rutas absolutas o relativas
BASE_DIR = Path(__file__).resolve().parent
RUTA_ORIGEN = BASE_DIR / INPUT_DIR / NOMBRE_ARCHIVO_ORIGEN
RUTA_SALIDA = BASE_DIR / OUTPUT_DIR / ARCHIVO_SALIDA

def main():
    print(f"Iniciando transformación desde: {RUTA_ORIGEN}")
    
    # Asegurar que el directorio output existe
    (BASE_DIR / OUTPUT_DIR).mkdir(exist_ok=True)

    # --- 1. LECTURA DE FECHA ---
    try:
        wb = load_workbook(filename=RUTA_ORIGEN, data_only=True)
        ws_fecha = wb[HOJA_FECHA]
        fecha = ws_fecha["J3"].value
        if isinstance(fecha, datetime):
            fecha = fecha.date()
    except FileNotFoundError:
        print(f"❌ ERROR: No se encuentra {NOMBRE_ARCHIVO_ORIGEN} en la carpeta input.")
        # En GitHub Actions, si falta el archivo, queremos que el proceso falle para avisarnos
        exit(1)
    except Exception as e:
        print(f"Advertencia leyendo fecha: {e}")
        fecha = np.nan
    
    # --- 2. LECTURA DE DATOS ---
    df = pd.read_excel(RUTA_ORIGEN, sheet_name=HOJA_DATOS, usecols="A:D", header=3)
    
    # Limpieza
    df.columns = df.columns.str.strip()
    df["Nombre proveedor"] = df["Nombre proveedor"].ffill().str.strip()
    df["Zona"] = df["Zona"].ffill().str.strip()
    df["Tipo Unidad"] = df["Tipo Unidad"].astype(str).str.strip()
    df["Kilos netos"] = pd.to_numeric(df["Kilos netos"], errors='coerce').fillna(0)
    
    # Filtrar
    df = df[df['Tipo Unidad'].isin(['Reclamo', 'Venta'])].copy()

    # --- 3. PIVOTADO ---
    df_pivot = df.pivot_table(
        index=['Nombre proveedor', 'Zona'],
        columns='Tipo Unidad',
        values='Kilos netos',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    df_pivot.columns.name = None
    df_pivot = df_pivot.rename_axis(columns=None)
    df_pivot.insert(0, "fecha", fecha)

    # --- 4. GUARDADO ---
    try:
        df_pivot.to_excel(RUTA_SALIDA, sheet_name='Hoja_Pivotada', index=False)
        print(f"✅ Archivo guardado correctamente en: {RUTA_SALIDA}")
    except Exception as e:
        print(f"❌ Error al guardar: {e}")
        exit(1)

if __name__ == "__main__":
    main()
