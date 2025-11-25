import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
import re

# --- CONFIGURACIÓN ---
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

NOMBRE_ARCHIVO_ORIGEN = "archivo.xlsx"
ARCHIVO_SALIDA = "datos_transformados.xlsx"
HOJA_DATOS = "RECLAM. FB"
HOJA_METADATOS = "RESULTADOS FB" # Usaremos esta hoja para FECHA y PRODUCTO

RUTA_ORIGEN = INPUT_DIR / NOMBRE_ARCHIVO_ORIGEN
RUTA_SALIDA = OUTPUT_DIR / ARCHIVO_SALIDA

def encontrar_producto_en_resultados(wb, hoja_nombre):
    """
    Busca el producto en la hoja 'RESULTADOS FB', celda C2.
    Texto esperado: 'RESULTADOS NETOS POR CLIENTES FRAMBUESA'
    """
    producto_detectado = "DESCONOCIDO"
    
    if hoja_nombre not in wb.sheetnames:
        return producto_detectado
        
    ws = wb[hoja_nombre]
    
    # Leemos la celda C2 directamente, como pediste
    celda_c2 = ws["C2"].value
    
    if celda_c2 and isinstance(celda_c2, str):
        texto = str(celda_c2).strip().upper()
        print(f"Texto encontrado en {hoja_nombre}!C2: '{texto}'")
        
        # Lógica de extracción:
        # Buscamos la palabra que viene después de "CLIENTES"
        # Ejemplo: "RESULTADOS NETOS POR CLIENTES FRAMBUESA" -> "FRAMBUESA"
        if "CLIENTES" in texto:
            partes = texto.split("CLIENTES")
            if len(partes) > 1:
                # Tomamos la parte derecha y limpiamos espacios
                # Quitamos posibles caracteres extra que no sean letras
                resto = partes[1].strip()
                # Cogemos la primera palabra que aparezca
                producto_detectado = resto.split()[0]
        else:
            # Si no dice "CLIENTES", intentamos coger la última palabra como plan B
            producto_detectado = texto.split()[-1]

    return producto_detectado

def main():
    print(f"--- Iniciando proceso en GitHub Actions ---")
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    if not RUTA_ORIGEN.exists():
        print(f"❌ ERROR: No se encuentra '{NOMBRE_ARCHIVO_ORIGEN}' en input.")
        exit(1)

    # Variables iniciales
    fecha = np.nan
    producto = "DESCONOCIDO"

    # --- 1. LECTURA DE METADATOS (FECHA Y PRODUCTO) ---
    try:
        wb = load_workbook(filename=RUTA_ORIGEN, data_only=True)
        
        # A) Leer FECHA de 'RESULTADOS FB' (J3)
        if HOJA_METADATOS in wb.sheetnames:
            ws_meta = wb[HOJA_METADATOS]
            val_fecha = ws_meta["J3"].value
            if isinstance(val_fecha, datetime):
                fecha = val_fecha.date()
            else:
                try:
                    fecha = datetime.strptime(str(val_fecha).split()[0], "%d/%m/%Y").date()
                except:
                    pass
        
        # B) Leer PRODUCTO de 'RESULTADOS FB' (C2)
        producto = encontrar_producto_en_resultados(wb, HOJA_METADATOS)
        print(f"✅ Producto detectado: {producto}")

    except Exception as e:
        print(f"Advertencia leyendo metadatos: {e}")

    # --- 2. LECTURA DE DATOS CON PANDAS ---
    print("Leyendo datos...")
    df = pd.read_excel(RUTA_ORIGEN, sheet_name=HOJA_DATOS, usecols="A:D", header=3)
    
    # Limpieza
    df.columns = df.columns.str.strip()
    df["Nombre proveedor"] = df["Nombre proveedor"].ffill().str.strip()
    df["Zona"] = df["Zona"].ffill().str.strip()
    df["Tipo Unidad"] = df["Tipo Unidad"].astype(str).str.strip()
    df["Kilos netos"] = pd.to_numeric(df["Kilos netos"], errors='coerce').fillna(0)
    
    # Filtrar
    df = df[df['Tipo Unidad'].isin(['Reclamo', 'Venta'])].copy()

    # --- 3. TRANSFORMACIÓN ---
    df_pivot = df.pivot_table(
        index=['Nombre proveedor', 'Zona'],
        columns='Tipo Unidad',
        values='Kilos netos',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    df_pivot.columns.name = None
    df_pivot = df_pivot.rename_axis(columns=None)
    
    # --- 4. AÑADIR COLUMNAS NUEVAS ---
    # Insertamos Fecha (col 0) y Producto (col 1)
    df_pivot.insert(0, "fecha", fecha)
    df_pivot.insert(1, "producto", producto)

    # --- 5. GUARDADO ---
    try:
        df_pivot.to_excel(RUTA_SALIDA, sheet_name='Hoja_Pivotada', index=False)
        print(f"✅ Archivo guardado correctamente en: {RUTA_SALIDA}")
    except Exception as e:
        print(f"❌ Error al guardar: {e}")
        exit(1)

if __name__ == "__main__":
    main()
