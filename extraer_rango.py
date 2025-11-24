import pandas as pd

# Archivo dentro de input/
archivo = "input/archivo.xlsx"

# Nombre de la hoja
hoja = "RECLAM. FB"

def main():
    # Leer SOLO el rango A4:E14
    df = pd.read_excel(
        archivo,
        sheet_name=hoja,
        usecols="A:E",
        skiprows=3,   # salta las primeras 3 filas → empieza en A4
        nrows=11      # 11 filas de A4 a A14
    )

    # Guardar resultado
    df.to_excel("output/rango_extraido.xlsx", index=False)
    print("✔ Rango A4:E14 extraído correctamente.")

if __name__ == "__main__":
    main()
