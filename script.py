import PyPDF2
import pandas as pd
import os

# Leer el contenido del PDF
def extract_text_from_pdf(file_path):
    pdf_reader = PyPDF2.PdfReader(file_path)
    text = ''
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text

# Parsear el texto extraído
def parse_invoice_text(text):
    lines = text.split('\n')
    data = {
        "SERIE": None,
        "BOLETA": None,
        "FECHA": None,
        "CLIENTE": None,
        "DNI": None,
        "PRODUCTOS": []
    }
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("EB"):
            data["SERIE"] = line.split('-')[0]
            data["BOLETA"] = line.split('-')[1]
        elif line.startswith("Fecha de Emisión"):
            data["FECHA"] = lines[i].replace("Fecha de Emisión :", "").strip()
            i += 1
        elif line.startswith("Señor(es)"):
            data["CLIENTE"] = lines[i + 1].strip()
            i += 1
        elif line.startswith("DNI"):
            data["DNI"] = lines[i + 1].strip()
            i += 1
        elif any(char.isdigit() for char in line) and "UNIDAD" in line:
            # Extraer los productos
            parts = line.split(' ')
            cantidad = parts[0].replace("UNIDAD", "").strip()
            unidad = parts[1].strip()
            descripcion = ' '.join(parts[2:len(parts)-3]).strip()
            valor_unitario = parts[-3].strip()
            data["PRODUCTOS"].append({
                "CANTIDAD": cantidad,
                "DESCRIPCIÓN": f"{unidad} {descripcion}",
                "VALOR_UNITARIO": valor_unitario
            })
        i += 1
    
    return data

# Guardar los datos en un archivo Excel
def save_to_excel(all_data, file_path):
    productos_df = pd.DataFrame(all_data)
    productos_df.to_excel(file_path, index=False)

# Obtener todos los archivos PDF en un directorio
def get_pdf_files(directory):
    return [os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.pdf')]

script_dir = os.path.dirname(os.path.abspath(__file__))
# Directorio de archivos PDF
pdf_directory = script_dir # Cambia esto a la ruta de tu directorio de PDFs
# Ruta del archivo Excel de salida
excel_path = 'boleta_consolidada.xlsx'

# Extraer y procesar todos los PDFs
all_productos_data = []
pdf_files = get_pdf_files(pdf_directory)

for pdf_path in pdf_files:
    pdf_text = extract_text_from_pdf(pdf_path)
    parsed_data = parse_invoice_text(pdf_text)
    productos_data = [{
        "SERIE": parsed_data["SERIE"],
        "BOLETA": parsed_data["BOLETA"],
        "FECHA": parsed_data["FECHA"],
        "CANTIDAD": producto["CANTIDAD"],
        "DESCRIPCIÓN": producto["DESCRIPCIÓN"],
        "VALOR_UNITARIO": producto["VALOR_UNITARIO"]
    } for producto in parsed_data["PRODUCTOS"]]
    all_productos_data.extend(productos_data)

# Guardar todos los datos en un archivo Excel
save_to_excel(all_productos_data, excel_path)

print("Datos extraídos y guardados en Excel con éxito.")
