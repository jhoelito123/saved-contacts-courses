import pandas as pd
import re
from pathlib import Path

ARCHIVO_EXCEL = "BaseDatosAdam.xlsx"
ARCHIVO_SALIDA = "contactos.vcf"
PREFIJO = "BD - "
CODIGO_PAIS = "+591"

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_numero(numero):
    if pd.isna(numero):
        return None

    numero = str(numero)

    digitos = re.sub(r"\D", "", numero)

    if len(digitos) == 0:
        return None

    if digitos.startswith("591"):
        telefono = "+" + digitos

    elif len(digitos) == 8:
        telefono = CODIGO_PAIS + digitos

    # Otro formato
    else:
        telefono = "+" + digitos

    # Validación mínima
    if len(re.sub(r"\D", "", telefono)) < 11:
        return None

    return telefono


def detectar_hoja_y_header(excel):

    for hoja in excel.sheet_names:

        # Leer únicamente las primeras filas
        preview = pd.read_excel(
            excel,
            sheet_name=hoja,
            header=None,
            nrows=20
        )

        for fila in range(len(preview)):

            valores = [
                str(x).upper().strip()
                for x in preview.iloc[fila].fillna("")
            ]

            if "CELULAR" in valores:
                return hoja, fila

    raise Exception("No se encontró la hoja con la columna CELULAR.")


def crear_vcard(nombre, telefono):

    return f"""BEGIN:VCARD
VERSION:3.0
FN:{nombre}
TEL;TYPE=CELL:{telefono}
END:VCARD
"""

# CARGAR EXCEL
excel = pd.ExcelFile(ARCHIVO_EXCEL)
hoja, fila_header = detectar_hoja_y_header(excel)

print("=" * 60)
print("Hoja encontrada :", hoja)
print("Encabezados en fila :", fila_header + 1)
print("=" * 60)

df = pd.read_excel(
    excel,
    sheet_name=hoja,
    header=fila_header
)

df.columns = [str(c).strip().upper() for c in df.columns]

# VALIDAR COLUMNAS
columnas = ["APE_PAT", "APE_MAT", "NOMBRES", "CELULAR"]

for c in columnas:
    if c not in df.columns:
        raise Exception(f"No existe la columna {c}")


vcards = []
telefonos = set()
procesados = 0
exportados = 0
ignorados = 0
duplicados = 0

for _, fila in df.iterrows():

    procesados += 1

    apellido_pat = limpiar_texto(fila["APE_PAT"])
    apellido_mat = limpiar_texto(fila["APE_MAT"])
    nombres = limpiar_texto(fila["NOMBRES"])

    telefono = limpiar_numero(fila["CELULAR"])

    if telefono is None:
        ignorados += 1
        continue

    if telefono in telefonos:
        duplicados += 1
        continue

    telefonos.add(telefono)

    nombre = " ".join([
        apellido_pat,
        apellido_mat,
        nombres
    ])

    nombre = re.sub(r"\s+", " ", nombre).strip()

    if nombre == "":
        ignorados += 1
        continue

    nombre = PREFIJO + nombre
    vcards.append(crear_vcard(nombre, telefono))
    exportados += 1

# GUARDAR
with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as archivo:
    archivo.write("\n".join(vcards))

# RESUMEN
print()
print("=" * 60)
print("PROCESO FINALIZADO")
print("=" * 60)

print(f"Archivo leído       : {ARCHIVO_EXCEL}")
print(f"Hoja utilizada      : {hoja}")
print(f"Filas Procesadas    : {procesados}")
print(f"Exportados          : {exportados}")
print(f"Ignorados           : {ignorados}")
print(f"Duplicados          : {duplicados}")
print(f"VCF generado        : {Path(ARCHIVO_SALIDA).resolve()}")

print("=" * 60)