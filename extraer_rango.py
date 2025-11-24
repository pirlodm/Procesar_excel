import pandas as pd

archivo = "input/archivo.xlsx"
hoja_fecha = "RESULTADOS FB"
hoja_datos = "RECLAM. FB"

def main():
    # 1) Leer la fecha de J3
    df_fecha = pd.read_excel(
        archivo,
        sheet_name=hoja_fecha,
        usecols="J",
        skiprows=2,  # fila 3
        nrows=1
    )
    fecha = df_fecha.iloc[0, 0]
    fecha = pd.to_datetime(fecha).date()  # solo fecha

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
