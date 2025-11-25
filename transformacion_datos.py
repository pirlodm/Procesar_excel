import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import re

# --- CONFIGURACIÓN ---
INPUT_DIR = "input"
OUTPUT_DIR = "output"
NOMBRE_ARCHIVO_ORIGEN = "archivo.xlsx"
ARCHIVO_SALIDA = "datos_transformados.xlsx"
HOJA_DATOS = "RECLAM. FB"
HOJA_FECHA = "RESULTADOS FB"

# Definir rutas
BASE_DIR = Path(__file__).resolve().parent
RUTA_ORIGEN = BASE_DIR / INPUT_DIR / NOMBRE_ARCHIVO_ORIGEN
RUTA_SALIDA = BASE_DIR / OUTPUT_DIR / ARCHIVO_SALIDA

def main():
    print(f"Iniciando transformación desde: {RUTA_ORIGEN}")
    
    # Asegurar que el directorio output existe
    (BASE_DIR / OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # Variables para almacenar lo que leamos con openpyxl
    fecha = np.nan
    producto = "DESCONOCIDO"

    # --- 1. LECTURA DE CABECERAS (FECHA Y PRODUCTO) ---
    try:
        wb = load_workbook(filename=RUTA_ORIGEN, data_only=True)
        
        # A) Leer FECHA de la hoja RESULTADOS FB
        if HOJA_FECHA in wb.sheetnames:
            ws_fecha = wb[HOJA_FECHA]
            val_fecha = ws_fecha["J3"].value
            if isinstance(val_fecha, datetime):
                fecha = val_fecha.date()
        
        # B) Leer PRODUCTO de la hoja RECLAM. FB (Fila 2, Celda C2)
        # Aunque esté combinada C-J, el valor vive en C2.
        if HOJA_DATOS in wb.sheetnames:
            ws_datos = wb[HOJA_DATOS]
            # Leemos la celda C2
            texto_cabecera = ws_datos["C2"].value
            
            if texto_cabecera:
                # Convertimos a texto, quitamos espacios y dividimos por palabras
                palabras = str(texto_cabecera).strip().split()
                if palabras:
                    # Tomamos la última palabra
                    producto = palabras[-1]
                    
                    # Opcional: Si quieres limpiar puntuación (puntos, comas) de esa palabra:
                    # producto = re.sub(r'[^\w\s]', '', producto)

    except FileNotFoundError:
        print(f"❌ ERROR: No se encuentra {NOMBRE_ARCHIVO_ORIGEN} en input.")
        exit(1)
    except Exception as e:
        print(f"Advertencia leyendo metadatos: {e}")

    # --- 2. LECTURA DE DATOS (PANDAS) ---
    print(f"Producto detectado: {producto}")
    
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
    
    # --- 4. INSERTAR COLUMNAS NUEVAS ---
    # Insertamos la fecha en la posición 0 (primera columna)
    df_pivot.insert(0, "fecha", fecha)
    
    # Insertamos el producto en la posición 1 (segunda columna)
    df_pivot.insert(1, "producto", producto)

    # --- 5. GUARDADO ---
    try:
        df_pivot.to_excel(RUTA_SALIDA, sheet_name='Hoja_Pivotada', index=False)
        print(f"✅ Archivo guardado correctamente en: {RUTA_SALIDA}")
    except Exception as e:
        print(f"❌ Error al guardar: {e}")

if __name__ == "__main__":
    main()
