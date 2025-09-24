import pandas as pd
from pathlib import Path
from datetime import datetime
import os
import requests
from datetime import datetime

hoy = datetime.now()
año = hoy.year - 2000
mes = hoy.month
trimestre_letra = ['a', 'b', 'c', 'd'][(mes - 1) // 3]
name_archivo = f"2_1_2{trimestre_letra}{año}_smc.xls"
url = f"https://www.bcv.org.ve/sites/default/files/EstadisticasGeneral/{name_archivo}"
nombre_archivo = os.path.basename(url)

response = requests.get(url, verify=False)
with open(nombre_archivo, "wb") as f:
    f.write(response.content)

print(f"Descarga completaa para {nombre_archivo}")

def extract_date(date_str):
    """Extrae la fecha en formato dd/mm/yyyy de un string."""
    if isinstance(date_str, str):
        import re
        match = re.search(r"(\d{2}/\d{2}/\d{4})", date_str)
        if match:
            try:
                return datetime.strptime(match.group(1), "%d/%m/%Y")
            except Exception:
                pass
    return None

def scrap_xls_to_txt(xls_path, txt_path):
    xls = pd.ExcelFile(xls_path)
    data = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        date_cell = df.iloc[4, 3] if df.shape[0] > 4 and df.shape[1] > 3 else None
        value_cell = df.iloc[14, 5] if df.shape[0] > 14 and df.shape[1] > 5 else None
        date = extract_date(str(date_cell))
        if date and value_cell is not None:
            data.append((date, value_cell))
        else:
            print(f"  ⚠ Datos no válidos en hoja '{sheet_name}': fecha extraída={date}, valor extraído={value_cell}")
    # Ordenar por fecha
    data.sort(key=lambda x: x[0])
    # Escribir en TXT
    with open(txt_path, "w", encoding="utf-8") as f:
        for date, value in data:
            f.write(f"{date.strftime('%d/%m/%Y')}: {value}\n")
    print(f"Datos extraídos y guardados en {txt_path}")

if __name__ == "__main__":
    folder = os.path.dirname(os.path.abspath(__file__))
    xls_files = [f for f in os.listdir(folder) if f.lower().endswith('.xls')]
    if not xls_files:
        print("No se encontraron archivos .xls en la carpeta actual.")
    for xls_file in xls_files:
        xls_path = os.path.join(folder, xls_file)
        txt_file = os.path.splitext(xls_file)[0] + ".txt"
        txt_path = os.path.join(folder, txt_file)
        print(f"\nProcesando archivo: {xls_file}")
        scrap_xls_to_txt(xls_path, txt_path)
