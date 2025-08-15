import os
import io
import re
import gc
import json
import tempfile
import cv2
import numpy as np
import pandas as pd
from pdf2image import pdfinfo_from_path, convert_from_path
from PIL import Image
from flask import Flask, request, render_template, send_file
from google.cloud import vision

# -----------------------
# Load Google Vision Credentials from ENV
# -----------------------
google_creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
if not google_creds_json:
    raise RuntimeError("Missing GOOGLE_APPLICATION_CREDENTIALS_JSON environment variable.")

# Create a temp credentials file for Google Vision API
with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as temp_cred_file:
    temp_cred_file.write(google_creds_json.encode())
    temp_cred_path = temp_cred_file.name

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = temp_cred_path

# -----------------------
# Flask app setup
# -----------------------
app = Flask(__name__, template_folder="templates", static_folder="static")

# Google Vision Client
client = vision.ImageAnnotatorClient()

# -----------------------
# OCR for one cell
# -----------------------
def ocr_cell_google(cell_img):
    img_byte_arr = io.BytesIO()
    Image.fromarray(cv2.cvtColor(cell_img, cv2.COLOR_BGR2RGB)).save(img_byte_arr, format='PNG')
    image = vision.Image(content=img_byte_arr.getvalue())
    response = client.document_text_detection(image=image)
    if response.error.message:
        raise Exception(f"OCR Error: {response.error.message}")
    return response.full_text_annotation.text

# -----------------------
# Extract structured data from text
# -----------------------
def extract_from_cell_text(text):
    voter_id_match = re.search(r"[A-Z]{3}\s*\d{7}", text)
    voter_id = voter_id_match.group(0).replace(" ", "") if voter_id_match else ""
    name_match = re.search(r"निर्वाचक का नाम[:\s]*([^\n]+)", text)
    name = name_match.group(1).strip() if name_match else ""
    rel_match = re.search(r"(?:पिता का नाम|पति का नाम|अन्य)[:\s]*([^\n]+)", text)
    relative = rel_match.group(1).strip() if rel_match else ""
    house_match = re.search(r"(?:मकान संख्या)[:\s]*([^\n]+)", text)
    house = house_match.group(1).strip() if house_match else ""
    age_match = re.search(r"उम्र[:\s]*([0-9]{1,3})", text)
    age = age_match.group(1).strip() if age_match else ""
    gender_match = re.search(r"लिंग[:\s]*([^\n]+)", text)
    gender = gender_match.group(1).strip() if gender_match else ""
    return [voter_id, name, relative, house, age, gender]

# -----------------------
# Process PDF with Google Vision
# -----------------------
def process_pdf_with_google(pdf_path, save_path):
    all_voters = []
    info = pdfinfo_from_path(pdf_path)
    total_pages = info['Pages']

    for page_number in range(1, total_pages + 1):
        pages = convert_from_path(pdf_path, dpi=190, first_page=page_number, last_page=page_number)
        page_img = pages[0]
        img = cv2.cvtColor(np.array(page_img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

        kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
        kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
        detect_h = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_h, iterations=2)
        detect_v = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel_v, iterations=2)
        grid = cv2.add(detect_h, detect_v)

        contours, _ = cv2.findContours(grid, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        boxes = [(x, y, w, h) for (x, y, w, h) in [cv2.boundingRect(c) for c in contours] if w > 200 and h > 200]

        unique_boxes = []
        for b in sorted(boxes, key=lambda b: (b[1], b[0])):
            if not any(abs(b[0] - ub[0]) < 15 and abs(b[1] - ub[1]) < 15 for ub in unique_boxes):
                unique_boxes.append(b)

        for (x, y, w, h) in unique_boxes:
            cell_img = img[y:y+h, x:x+w]
            cell_text = ocr_cell_google(cell_img)
            voter_data = extract_from_cell_text(cell_text)
            if any(voter_data):
                all_voters.append(voter_data)
            del cell_img, cell_text, voter_data
            gc.collect()

        del img, gray, thresh, detect_h, detect_v, grid, contours, boxes, unique_boxes, page_img, pages
        gc.collect()

    df = pd.DataFrame(all_voters, columns=["VoterID", "VoterName", "RelativeName", "HouseNumber", "Age", "Gender"])
    with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Voters")
        ws = writer.sheets["Voters"]
        ws.set_column("A:A", 15)
        ws.set_column("B:B", 25)
        ws.set_column("C:C", 25)
        ws.set_column("D:D", 15)
        ws.set_column("E:E", 5)
        ws.set_column("F:F", 10)

# -----------------------
# Routes
# -----------------------
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "pdf_file" not in request.files:
        return "No file uploaded", 400

    pdf_file = request.files["pdf_file"]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        pdf_file.save(temp_pdf.name)
        pdf_path = temp_pdf.name

    output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name

    try:
        process_pdf_with_google(pdf_path, output_path)
        return send_file(output_path, as_attachment=True, download_name="output.xlsx")
    except Exception as e:
        print("Error:", str(e))
        return {"error": str(e)}, 500

# -----------------------
# Run app (Railway sets PORT)
# -----------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
