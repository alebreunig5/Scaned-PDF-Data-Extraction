# Funcion para no analizar mediante OCR a PDFs que ya tengan su archivo .txt correspondiente
# Actualización de la API de Gemini
# API en archivo .venv
# Print para cada proceso

import os
import pytesseract
from pdf2image import convert_from_path
import openpyxl
import json
import google.genai as genai
import time
from dotenv import load_dotenv 


# --- CONFIGURACIÓN ---
PDF_FOLDER = 'PDFs'
EXCEL_OUTPUT_FILE = 'datos_extraidos_gemini.xlsx'
OCR_OUTPUT_FOLDER = 'text_output'

# --- CARGAR CONFIGURACIÓN SEGURA ---
load_dotenv() 
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") 
GEMINI_MODEL = 'gemini-2.0-flash-lite' # O 'gemini-1.5-flash-lite' para ahorrar

# --- RUTAS ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\abreunig\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
poppler_path = r'C:\Users\abreunig\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin'

# --- FUNCIÓN 1: OCR de PDF (solo si no existe txt) ---
def extract_text_from_scanned_pdf(pdf_path, poppler_path=None):
    print(f"\n[1/4] Extrayendo texto de: {os.path.basename(pdf_path)}")
    try:
        pages = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
        full_text = ""
        for page_num, page_image in enumerate(pages):
            print(f"   Página {page_num + 1}...")
            text = pytesseract.image_to_string(page_image, lang='spa')
            full_text += text + "\n\n"
        print("   ✓ Texto extraído")
        return full_text
    except Exception as e:
        print(f"   ✗ Error OCR: {e}")
        return None

# --- FUNCIÓN 2: Guardar texto ---
def save_text_to_file(text_content, pdf_filename, output_folder):
    txt_filename = os.path.splitext(pdf_filename)[0] + '.txt'
    output_path = os.path.join(output_folder, txt_filename)
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text_content)
        print(f"[2/4] Texto guardado en: {txt_filename}")
    except Exception as e:
        print(f"   ✗ Error guardando texto: {e}")

# --- FUNCIÓN NUEVA: Leer texto existente ---
def read_existing_text(pdf_filename, output_folder):
    txt_filename = os.path.splitext(pdf_filename)[0] + '.txt'
    output_path = os.path.join(output_folder, txt_filename)
    
    if os.path.exists(output_path):
        print(f"📁 Leyendo texto existente de: {txt_filename}")
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
            if len(content.strip()) > 50:  # Verifica que tenga contenido útil
                print(f"   ✓ Texto cargado ({len(content)} caracteres)")
                return content
            else:
                print(f"   ⚠️  Archivo vacío o muy corto, reprocesando...")
                return None
        except Exception as e:
            print(f"   ✗ Error leyendo archivo: {e}, reprocesando...")
            return None
    return None

# --- FUNCIÓN 3: Gemini Analysis ---
def analyze_text_with_gemini(text_content, client, model_name):
    print("[3/4] Analizando con Gemini...")
    
    # Prompt optimizado
    prompt = f"""Extrae los siguientes datos del formulario y devuélvelos como JSON. Si un campo no existe, usa "N/A".

TEXTO DEL FORMULARIO:
{text_content}

INSTRUCCIONES:
- Fechas: formato DD/MM/YYYY
- Números: convertir a float (ej: 1.000,50 → 1000.5)
- Nombre: separar en "Apellido" y "Nombre"
- Booleanos: true/false según corresponda
- Mantén strings exactos como aparecen

DEVUELVE SOLO JSON con esta estructura:
{{
  "Apellido": "string",
  "Nombre": "string",
  "Nacionalidad": "string",
  "Tipo_Documento": "string",
  "Numero_Documento": "string",
  "Calle": "string",
  "Nro": "string",
  "Piso": "string",
  "Departamento": "string",
  "Localidad": "string",
  "Codigo_Postal": "string",
  "Provincia": "string",
  "Pais": "string",
  "Fecha_de_Nacimiento": "string",
  "Email": "string",
  "Radicada_en_el_Exterior": "boolean",
  "Radicada_en_Paraiso_Fiscal": "boolean",
  "Es_Peps": "boolean",
  "Fecha_de_Operacion": "string",
  "Tipo_de_Moneda": "string",
  "Monto_Total": "float",
  "Monto_Total_en_Pesos": "float",
  "Pago_en_favor_de_Terceros": "boolean",
  "Forma_de_Pago": "string",
  "Porcentaje_del_pago_total": "float",
  "Fecha_de_pago": "string",
  "REPORTE UIF": "string",
  "XLSM": "string",
  "cuil": "string",
  "N° DE FORMULARIO PREMIO": "string",
  "ACTIVIDAD DECLARADA": "string",
  "EMPRESA EN LA QUE DESARROLLA SUS TAREAS": "string",
  "ORIGEN DE FONDOS DECLARADO": "string",
  "PEP (SI/NO)": "string",
  "SO (SI/NO)": "string",
  "OBSERVACIONES FORMULARIO DE PAGO DE PREMIO": "string"
}}"""
    
    try:
        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config={"response_mime_type": "application/json"}
        )
        
        extracted_info = json.loads(response.text)
        print("   ✓ Análisis completado")
        
        # Limpiar datos
        for key in extracted_info:
            if isinstance(extracted_info[key], str):
                extracted_info[key] = extracted_info[key].strip()
                if extracted_info[key] == "":
                    extracted_info[key] = "N/A"
        
        # Mostrar resumen
        print("\n   RESUMEN:")
        importantes = ['Apellido', 'Nombre', 'Numero_Documento', 'Monto_Total']
        for key in importantes:
            if key in extracted_info:
                print(f"   • {key}: {extracted_info[key]}")
        
        return extracted_info
        
    except Exception as e:
        print(f"   ✗ Error Gemini: {e}")
        return None

# --- FUNCIÓN 4: Guardar en Excel ---
def save_to_excel(data_list, output_file):
    print(f"[4/4] Guardando en Excel: {output_file}")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos"
    
    # Encabezados
    headers = [
        'Apellido', 'Nombre', 'Nacionalidad', 'Tipo_Documento', 'Numero_Documento',
        'Calle', 'Nro', 'Piso', 'Departamento', 'Localidad', 'Codigo_Postal',
        'Provincia', 'Pais', 'Fecha_de_Nacimiento', 'Email',
        'Radicada_en_el_Exterior', 'Radicada_en_Paraiso_Fiscal', 'Es_Peps',
        'Fecha_de_Operacion', 'Tipo_de_Moneda', 'Monto_Total', 'Monto_Total_en_Pesos',
        'Pago_en_favor_de_Terceros', 'Forma_de_Pago', 'Porcentaje_del_pago_total',
        'Fecha_de_pago', 'REPORTE UIF', 'XLSM', 'cuil', 'N° DE FORMULARIO PREMIO',
        'ACTIVIDAD DECLARADA', 'EMPRESA EN LA QUE DESARROLLA SUS TAREAS',
        'ORIGEN DE FONDOS DECLARADO', 'PEP (SI/NO)', 'SO (SI/NO)',
        'OBSERVACIONES FORMULARIO DE PAGO DE PREMIO', 'Archivo_Origen'
    ]
    
    ws.append(headers)
    
    # Datos
    for data in data_list:
        row = [data.get(h, 'N/A') for h in headers[:-1]]
        row.append(data.get('archivo', 'N/A'))
        ws.append(row)
    
    wb.save(output_file)
    print(f"   ✓ Excel guardado con {len(data_list)} registros")

# --- FUNCIÓN PRINCIPAL MEJORADA ---
def main():
    print("=" * 60)
    print("EXTRACTOR DE DATOS DE FORMULARIOS PDF (OPTIMIZADO)")
    print("=" * 60)
    
    # 1. Validar API Key
    if not GEMINI_API_KEY or GEMINI_API_KEY == "TU_CLAVE_AQUÍ":
        print("\n⚠️  ERROR: Configura tu API Key de Gemini")
        print("   Obtén una en: https://aistudio.google.com/app/apikey")
        return
    
    # 2. Inicializar Gemini
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        print("✓ Cliente Gemini inicializado")
    except Exception as e:
        print(f"✗ Error inicializando Gemini: {e}")
        return
    
    # 3. Verificar carpetas
    if not os.path.isdir(PDF_FOLDER):
        print(f"✗ Carpeta '{PDF_FOLDER}' no encontrada")
        return
    
    # Crear carpeta text_output si no existe
    if not os.path.exists(OCR_OUTPUT_FOLDER):
        os.makedirs(OCR_OUTPUT_FOLDER)
        print(f"✓ Carpeta '{OCR_OUTPUT_FOLDER}' creada")
    
    # 4. Procesar PDFs (CON LA NUEVA LÓGICA)
    all_data = []
    pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"✗ No hay PDFs en la carpeta '{PDF_FOLDER}'")
        return
    
    print(f"\n📁 Encontrados {len(pdf_files)} PDF(s) para procesar")
    
    # Verificar archivos de texto existentes
    existing_txt_files = [f for f in os.listdir(OCR_OUTPUT_FOLDER) if f.endswith('.txt')]
    print(f"📄 Archivos .txt encontrados en '{OCR_OUTPUT_FOLDER}': {len(existing_txt_files)}")
    
    for i, filename in enumerate(pdf_files, 1):
        print(f"\n{'─' * 40}")
        print(f"📄 [{i}/{len(pdf_files)}] {filename}")
        print(f"{'─' * 40}")
        
        pdf_path = os.path.join(PDF_FOLDER, filename)
        text = None
        
        # ✅ NUEVA LÓGICA: Verificar si ya existe el texto extraído
        existing_text = read_existing_text(filename, OCR_OUTPUT_FOLDER)
        
        if existing_text:
            # Usar texto existente
            text = existing_text
            print("✅ Usando texto previamente extraído (se omite OCR)")
        else:
            # Extraer texto del PDF (OCR)
            text = extract_text_from_scanned_pdf(pdf_path, poppler_path)
            
            if not text or len(text.strip()) < 50:
                print("   ✗ Texto insuficiente o error en OCR")
                continue
            
            # Guardar texto para futuras ejecuciones
            save_text_to_file(text, filename, OCR_OUTPUT_FOLDER)
        
        # Analizar con Gemini (tanto para texto nuevo como existente)
        data = analyze_text_with_gemini(text, client, GEMINI_MODEL)
        
        if data:
            data['archivo'] = filename
            all_data.append(data)
        
        # Pausa entre requests (evitar rate limits)
        if i < len(pdf_files):
            print("   ⏳ Pausa 5 segundos...")
            time.sleep(5)
    
    # 5. Guardar resultados
    if all_data:
        save_to_excel(all_data, EXCEL_OUTPUT_FILE)
        
        # Resumen final
        print(f"\n{'=' * 60}")
        print("✅ PROCESO COMPLETADO")
        print(f"{'=' * 60}")
        print(f"• PDFs procesados: {len(all_data)}/{len(pdf_files)}")
        print(f"• Textos guardados/reutilizados en: {OCR_OUTPUT_FOLDER}/")
        print(f"• Excel generado: {EXCEL_OUTPUT_FILE}")
        
        # Mostrar estadísticas de reutilización
        if existing_txt_files:
            reused = len([f for f in pdf_files 
                         if os.path.exists(os.path.join(OCR_OUTPUT_FOLDER, 
                                                       os.path.splitext(f)[0] + '.txt'))])
            print(f"• Archivos reutilizados: {reused} (sin reprocesar OCR)")
        
        # Mostrar ubicación completa
        excel_path = os.path.abspath(EXCEL_OUTPUT_FILE)
        print(f"• Ruta completa: {excel_path}")
    else:
        print("\n⚠️  No se extrajeron datos válidos de ningún PDF")

# --- EJECUCIÓN ---
if __name__ == "__main__":
    main()
