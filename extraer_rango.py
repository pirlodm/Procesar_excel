import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

archivo = "input/archivo.xlsx"
hoja_fecha = "RESULTADOS FB"
hoja_datos = "RECLAM. FB"

def main():
    # 1) Leer la celda J3 usando openpyxl
    wb = load_workbook(filename=archivo, data_only=True)
    ws_fecha = wb[hoja_fecha]
    valor = ws_fecha["J3"].value  # puede ser string

    # 2) Convertir a fecha
    if isinstance(valor, str):
        # Supongamos formato "DD/MM/YYYY HH:MM:SS"
        fecha = datetime.strptime(valor.split()[0], "%d/%m/%Y").date()
    else:
        # si ya es datetime
        fecha = valor.date() if hasattr(valor, "date") else valor

    # 3) Leer el rango de datos ajustado (A:D, 8 filas)
    df_datos = pd.read_excel(
        archivo,
        sheet_name=hoja_datos,
        usecols="A:D",
        skiprows=3,
        nrows=8
    )

    # 4) Añadir la fecha como primera columna
    df_datos.insert(0, "Fecha", fecha)

    # 5) Guardar resultado
    df_datos.to_excel("output/rango_extraido.xlsx", index=False)
    print("✔ Datos extraídos con fecha como primera columna, 3 filas menos y sin columna E.")

if __name__ == "__main__":
    main()
