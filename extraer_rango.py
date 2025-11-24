import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

archivo = "input/archivo.xlsx"
hoja_fecha = "RESULTADOS FB"
hoja_datos = "RECLAM. FB"

def main():
    # 1) Abrir workbook con openpyxl
    wb = load_workbook(filename=archivo, data_only=True)

    # 2) Leer fecha de J3 de RESULTADOS FB
    ws_fecha = wb[hoja_fecha]
    valor_fecha = ws_fecha["J3"].value
    if isinstance(valor_fecha, str):
        fecha = datetime.strptime(valor_fecha.split()[0], "%d/%m/%Y").date()
    else:
        fecha = valor_fecha.date() if hasattr(valor_fecha, "date") else valor_fecha

    # 3) Leer rango de datos de RECLAM. FB usando pandas
    df_datos = pd.read_excel(
        archivo,
        sheet_name=hoja_datos,
        usecols="A:D",
        skiprows=3,
        nrows=8
    )

    # 4) Repetir valor de B2 en toda la columna B
    ws_datos = wb[hoja_datos]
    valor_b2 = ws_datos["B2"].value
    df_datos["B"] = valor_b2

    # 5) Añadir la fecha como primera columna
    df_datos.insert(0, "Fecha", fecha)

    # 6) Guardar resultado
    df_datos.to_excel("output/rango_extraido.xlsx", index=False)
    print("✔ Datos extraídos con fecha en primera columna y B2 repetido en la columna B.")

if __name__ == "__main__":
    main()
