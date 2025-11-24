import pandas as pd
from openpyxl import load_workbook

archivo = "input/archivo.xlsx"
hoja_fecha = "RESULTADOS FB"
hoja_datos = "RECLAM. FB"

def main():
    # 1) Leer la fecha de J3 usando openpyxl
    wb = load_workbook(filename=archivo, data_only=True)
    ws_fecha = wb[hoja_fecha]
    fecha = ws_fecha["J3"].value  # lee directamente J3

    # Opcional: convertir a solo fecha si tiene hora
    if hasattr(fecha, "date"):
        fecha = fecha.date()

    # 2) Leer el rango de datos ajustado (A:D, 8 filas)
    df_datos = pd.read_excel(
        archivo,
        sheet_name=hoja_datos,
        usecols="A:D",
        skiprows=3,
        nrows=8
    )

    # 3) Añadir la fecha como primera columna
    df_datos.insert(0, "Fecha", fecha)

    # 4) Guardar resultado
    df_datos.to_excel("output/rango_extraido.xlsx", index=False)
    print("✔ Datos extraídos con fecha (J3) como primera columna, 3 filas menos y sin columna E.")

if __name__ == "__main__":
    main()
