from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
from pdfminer.high_level import extract_text
from decimal import Decimal
import openpyxl
import re
from collections import defaultdict
import io
import os
from fastapi.middleware.cors import CORSMiddleware
import uuid





# Your existing script (with slight modifications if necessary)
def findLine(text , search):
    for lineIndex in range(len(text.split("\n"))):
        line = text.split("\n")[lineIndex]
        if search in line:
            return lineIndex


def normalize_key(key: str) -> str:
    normalized_key = key.lower()
    normalized_key = re.sub(r'[^a-z0-9]', '', normalized_key)
    return normalized_key


def extract(pdf_file_path):
    extractedData = defaultdict(list)
    with open(pdf_file_path, 'rb') as fh:
        for page_text in extract_text(fh).split("\f")[:-1]:
            sampleName = page_text.split("\n")[findLine(page_text , "Sample Name")].split(" ")
            sampleName = [x for x in sampleName if x]
            sampleName = normalize_key(sampleName[2])
            areaLine = page_text.split("\n")[findLine(page_text , "RetTime") + 3].split(" ")
            areaLine = [x for x in areaLine if x]
            Area = areaLine[4]
            extractedData[sampleName].append(Area)
    return extractedData


def push(extractedData, excel_template):
    dic = {
        "std50": [6, 3],
        "std80": [6, 4],
        "std100": [6, 5],
        "std160": [6, 6],
        "std200": [6, 7],
        "t80": [17, 4],
        "t100": [17, 5],
        "t160": [17, 6],
        "f1": [30, 7],
        "f2": [30, 9],
        "sday": [54, 8],
        "sanalyst": [30, 4],
        "scolumn": [43, 9],
        "m2": [42, 5],
        "m1": [42, 3],
        "stability": [54, 2]
    }

    workbook = openpyxl.load_workbook(excel_template)
    sheet = workbook.active

    for key, values in extractedData.items():
        if key not in dic:
            print(f"Key {key} not found in dictionary, skipping.")
            continue

        for value in values:
            print(f"Writing {key} value {value} to row {dic[key][0]} and column {dic[key][1]}")
            try:
                sheet.cell(row=dic[key][0], column=dic[key][1]).value = Decimal(value)
                dic[key][0] += 1
            except Exception as e:
                print(f"Error writing to Excel for {key}: {e}")

    result_filename = 'result.xlsx'
    workbook.save(result_filename)
    return result_filename


app = FastAPI()


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # You can replace '*' with specific domains if needed
    allow_credentials=True,
    allow_methods=["*"],  # Allow all HTTP methods (GET, POST, etc.)
    allow_headers=["*"],  # Allow all headers
)


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    # Temporary file to store the uploaded PDF
    pdf_path = f"temp_{uuid.uuid4().hex}.pdf"

    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    try:
        # Extract data from the uploaded PDF
        extracted_data = extract(pdf_path)

        # Process the extracted data and push it into the Excel template
        result_file_name = push(extracted_data, "template2.xlsx")

        # Return the result Excel file as a response
        return FileResponse(result_file_name, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='result.xlsx')

    finally:
        # Clean up the temporary PDF file
        os.remove(pdf_path)
