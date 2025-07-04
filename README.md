# Scaned-PDF-Data-Extraction
This project provides a Python script to extract structured data from scanned PDF documents using Optical Character Recognition (OCR) and the Google Gemini API. The extracted information is then saved into an Excel spreadsheet. This solution is particularly useful for automating data entry from forms or reports.

# Features
- OCR Integration: Utilizes pytesseract and pdf2image to extract text from scanned PDFs.

- Google Gemini API: Leverages the power of Gemini 1.5 Flash (or Pro) to intelligently parse and structure extracted text into a JSON format based on a predefined schema.

- Excel Output: Organizes all extracted data into a single, clean Excel (.xlsx) file.

- Error Handling: Includes basic error handling for file operations, API calls, and JSON parsing.

## How It Works
The script follows a sequential process for each PDF file found in a specified input folder:

1. PDF to Image Conversion: 
```pdf2image``` converts each page of the PDF into an image.

2. OCR Text Extraction: 
```pytesseract``` applies OCR to these images to extract raw text content.

3. Text File Storage: 
The extracted raw text for each PDF is saved into a .txt file for review and debugging.

4. Gemini API Analysis: 
The raw text is sent to the Google Gemini API with a detailed prompt and a JSON schema. Gemini then extracts the relevant data and returns it in a structured JSON format.

5. Data Structuring and Cleaning: 
The script performs post-processing on the data received from Gemini, such as converting monetary values to floats.

6. Excel Output: 
Finally, all the structured data from processed PDFs is compiled and saved into a single Excel file, with each data point mapped to its corresponding column.

## Setup and Installation (Windows)
To get this script running on your Windows machine, follow these steps:

### 1. Python Installation
If you don't have Python installed, download the latest version from the official Python website. During installation, make sure to check the box that says "Add Python to PATH".

### 2. Install Required Libraries
Open your Command Prompt (CMD) or PowerShell and run the following commands to install the necessary Python libraries:

```python
pip install pytesseract pdf2image openpyxl google-generativeai
```
### 3. Install Tesseract OCR
Tesseract is an open-source OCR engine. Download the Windows installer from the UB Mannheim Tesseract page.

- Run the installer.

- During installation, make sure to check the option to "Add to PATH" for all users. If you don't, you'll need to manually specify the Tesseract executable path in the script.

- Remember to install the language data for Spanish (spa) during the Tesseract setup, as the script uses it.

### 4. Install Poppler
```pdf2image``` relies on Poppler, a PDF rendering library. Download the pre-compiled binaries for Windows from the Poppler for Windows website. Look for a release like ```poppler-*-win64.zip```.

Unzip the downloaded file to a convenient location, for example, ```C:\Program Files\poppler-24.08.0```.

Note the path to the ```bin``` directory inside the unzipped folder (e.g., ```C:\Program Files\poppler-24.08.0\Library\bin```). You will need to configure this path in the script.

### 5. Get Your Google Gemini API Key
To use the Gemini API, you need an API key.

1. Go to Google AI Studio.

2. Sign in with your Google account.

3. Click on "Create API key in new project" or "Get API key" if you have an existing project.

4. Copy the generated API key.

Important: Never share your API key publicly or commit it directly to your version control (like Git). In this script, we'll store it directly, but for production environments, consider using environment variables or a secure configuration management system.

### 6. Configure the Script
Open the ```main.py``` script (or whatever you name your Python file) in a text editor and make the following adjustments:

- Gemini API Key: Replace ```"YOUR_API_KEY"``` with your actual Gemini API key.

```python
# --- CONFIGURACIÓN DE LA API DE GEMINI ---
# ¡IMPORTANTE! Reemplaza "TU_CLAVE_DE_API" con tu clave real.
GEMINI_API_KEY = "YOUR_API_KEY" # <--- PASTE YOUR API KEY HERE
genai.configure(api_key=GEMINI_API_KEY)
```
- Poppler Path: Uncomment and update the "```poppler_path"``` variable with the path to the bin directory of your Poppler installation (from step 4).

```python
# --- OPCIONAL: Configura las rutas de Tesseract y Poppler ---
# Si NO agregaste Tesseract y Poppler a las variables de entorno,
# descomenta y ajusta las siguientes líneas con las rutas de los ejecutables.
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' # Only if not added to PATH
poppler_path = r'C:\Program Files\poppler-24.08.0\Library\bin' # <--- UPDATE THIS PATH
```

### 7. Prepare Your PDFs
Create a folder named PDFs in the same directory as your Python script. Place all the scanned PDF documents you want to process into this PDFs folder.

## Running the Script
Once everything is set up, open your Command Prompt or PowerShell, navigate to the directory where you saved the script, and run it:

```python
python your_script_name.py
```
The script will print progress messages to the console. Upon successful completion, an Excel file named ```datos_extraidos_gemini.xlsx``` will be created in the same directory, containing all the extracted data. A new ```text_output``` folder will also be created with ```.txt``` files containing the raw OCR output for each PDF.


