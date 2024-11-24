import streamlit as st
import os
from tempfile import NamedTemporaryFile
from fpdf import FPDF
from PIL import Image
import markdown
import json
import xml.etree.ElementTree as ET
from ebooklib import epub
import pyttsx3
import pyqrcode
import barcode
from barcode.writer import ImageWriter
import wave
import speech_recognition as sr
from pptx import Presentation
import zipfile
import shutil
import pandas as pd
from docx import Document
import csv
import xlrd
from io import BytesIO

# Conversion Functions (Placeholder)

def convert_dwg_to_pdf(input_file, output_file):
    try:
        # Placeholder conversion logic
        return f"DWG to PDF conversion: {output_file}"
    except Exception as e:
        return f"Error in DWG to PDF conversion: {str(e)}"

def convert_rvt_to_dwg(input_file, output_file):
    try:
        # Placeholder conversion logic
        return f"RVT to DWG conversion: {output_file}"
    except Exception as e:
        return f"Error in RVT to DWG conversion: {str(e)}"

def convert_ppt_to_pdf(input_file, output_file):
    try:
        ppt = Presentation(input_file)
        pdf_writer = FPDF()
        pdf_writer.add_page()
        for slide in ppt.slides:
            pdf_writer.cell(200, 10, txt=slide.shapes.title.text, ln=True)
        pdf_writer.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in PPT to PDF conversion: {str(e)}"

def convert_txt_to_pdf(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            text = file.read()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in TXT to PDF conversion: {str(e)}"

def convert_md_to_pdf(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            text = markdown.markdown(file.read())
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, text)
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in Markdown to PDF conversion: {str(e)}"

def convert_html_to_pdf(input_file, output_file):
    try:
        html_content = open(input_file, 'r').read()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, html_content)
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in HTML to PDF conversion: {str(e)}"

def convert_epub_to_pdf(input_file, output_file):
    try:
        book = epub.read_epub(input_file)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                pdf.multi_cell(0, 10, item.content.decode())
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in EPUB to PDF conversion: {str(e)}"

def convert_image_to_pdf(input_file, output_file):
    try:
        image = Image.open(input_file)
        image.convert('RGB').save(output_file, "PDF")
        return output_file
    except Exception as e:
        return f"Error in Image to PDF conversion: {str(e)}"

def convert_audio_to_text(input_file):
    try:
        recognizer = sr.Recognizer()
        audio = sr.AudioFile(input_file)
        with audio as source:
            audio_data = recognizer.record(source)
        return recognizer.recognize_google(audio_data)
    except Exception as e:
        return f"Error in audio to text conversion: {str(e)}"

def convert_video_to_audio(input_file, output_audio):
    try:
        video_clip = VideoFileClip(input_file)
        video_clip.audio.write_audiofile(output_audio)
        return output_audio
    except Exception as e:
        return f"Error in video to audio conversion: {str(e)}"

def generate_qr_code(data, output_file):
    try:
        qr = pyqrcode.create(data)
        qr.png(output_file, scale=6)
        return output_file
    except Exception as e:
        return f"Error generating QR code: {str(e)}"

def generate_barcode(data, output_file):
    try:
        barcode_obj = barcode.get('ean13', data, writer=ImageWriter())
        barcode_obj.save(output_file)
        return output_file
    except Exception as e:
        return f"Error generating barcode: {str(e)}"

def html_to_text(input_file, output_file):
    try:
        with open(input_file, 'r') as f:
            html_content = f.read()
        text = html.unescape(html_content)
        with open(output_file, 'w') as f:
            f.write(text)
        return output_file
    except Exception as e:
        return f"Error in HTML to text conversion: {str(e)}"

# Additional Conversion Functions

def convert_docx_to_pdf(input_file, output_file):
    try:
        doc = Document(input_file)
        pdf = FPDF()
        pdf.add_page()
        for para in doc.paragraphs:
            pdf.multi_cell(0, 10, para.text)
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in DOCX to PDF conversion: {str(e)}"

def convert_xlsx_to_pdf(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        pdf = FPDF()
        pdf.add_page()
        for index, row in df.iterrows():
            pdf.multi_cell(0, 10, str(row.tolist()))
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in XLSX to PDF conversion: {str(e)}"

def convert_csv_to_pdf(input_file, output_file):
    try:
        df = pd.read_csv(input_file)
        pdf = FPDF()
        pdf.add_page()
        for index, row in df.iterrows():
            pdf.multi_cell(0, 10, str(row.tolist()))
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in CSV to PDF conversion: {str(e)}"

def convert_json_to_pdf(input_file, output_file):
    try:
        with open(input_file, 'r') as f:
            data = json.load(f)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, json.dumps(data, indent=4))
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in JSON to PDF conversion: {str(e)}"

def convert_xml_to_pdf(input_file, output_file):
    try:
        tree = ET.parse(input_file)
        root = tree.getroot()
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, ET.tostring(root, encoding='unicode'))
        pdf.output(output_file)
        return output_file
    except Exception as e:
        return f"Error in XML to PDF conversion: {str(e)}"

def convert_txt_to_csv(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            lines = file.readlines()
        with open(output_file, 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            for line in lines:
                writer.writerow([line.strip()])
        return output_file
    except Exception as e:
        return f"Error in TXT to CSV conversion: {str(e)}"

def compress_folder(input_folder, output_file):
    try:
        shutil.make_archive(output_file, 'zip', input_folder)
        return output_file + ".zip"
    except Exception as e:
        return f"Error in compressing folder: {str(e)}"

def extract_zip(input_file, output_folder):
    try:
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(output_folder)
        return f"Extracted zip to {output_folder}"
    except Exception as e:
        return f"Error in extracting ZIP: {str(e)}"

# Streamlit App UI Setup
st.title("Advanced File Conversion Platform")
st.write("Convert various file formats using this tool.")

# File upload section
uploaded_file = st.file_uploader("Upload your file", type=["dwg", "rvt", "pptx", "txt", "zip", "jpg", "jpeg", "png", "md", "html", "epub", "json", "csv", "xml", "docx", "xlsx", "mp4", "mp3"])

# Conversion options
conversion_type = st.selectbox(
    "Select the conversion format",
    [
        "DWG to PDF", "RVT to DWG", "PPT to PDF", "TXT to PDF", "MD to PDF", "HTML to PDF", 
        "EPUB to PDF", "Image to PDF", "Audio to Text", "Video to Audio", "QR Code Generator", 
        "Barcode Generator", "HTML to Text", "DOCX to PDF", "XLSX to PDF", "CSV to PDF", 
        "JSON to PDF", "XML to PDF", "TXT to CSV", "Compress Folder", "Extract ZIP"
    ]
)

# Perform file conversion and display result
if uploaded_file is not None and st.button("Convert"):
    try:
        with NamedTemporaryFile(delete=False) as tmp_file:
            tmp_file.write(uploaded_file.read())
            tmp_file_path = tmp_file.name

        if conversion_type == "DWG to PDF":
            output_file = tmp_file_path.replace(".dwg", ".pdf")
            result = convert_dwg_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "RVT to DWG":
            output_file = tmp_file_path.replace(".rvt", ".dwg")
            result = convert_rvt_to_dwg(tmp_file_path, output_file)
        elif conversion_type == "PPT to PDF":
            output_file = tmp_file_path.replace(".pptx", ".pdf")
            result = convert_ppt_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "TXT to PDF":
            output_file = tmp_file_path.replace(".txt", ".pdf")
            result = convert_txt_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "MD to PDF":
            output_file = tmp_file_path.replace(".md", ".pdf")
            result = convert_md_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "HTML to PDF":
            output_file = tmp_file_path.replace(".html", ".pdf")
            result = convert_html_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "EPUB to PDF":
            output_file = tmp_file_path.replace(".epub", ".pdf")
            result = convert_epub_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "Image to PDF":
            output_file = tmp_file_path.replace(".jpg", ".pdf").replace(".jpeg", ".pdf").replace(".png", ".pdf")
            result = convert_image_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "Audio to Text":
            result = convert_audio_to_text(tmp_file_path)
        elif conversion_type == "Video to Audio":
            output_audio = tmp_file_path.replace(".mp4", ".mp3")
            result = convert_video_to_audio(tmp_file_path, output_audio)
        elif conversion_type == "QR Code Generator":
            output_file = tmp_file_path.replace(".txt", ".png")
            result = generate_qr_code(tmp_file_path, output_file)
        elif conversion_type == "Barcode Generator":
            output_file = tmp_file_path.replace(".txt", ".png")
            result = generate_barcode(tmp_file_path, output_file)
        elif conversion_type == "HTML to Text":
            output_file = tmp_file_path.replace(".html", ".txt")
            result = html_to_text(tmp_file_path, output_file)
        elif conversion_type == "DOCX to PDF":
            output_file = tmp_file_path.replace(".docx", ".pdf")
            result = convert_docx_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "XLSX to PDF":
            output_file = tmp_file_path.replace(".xlsx", ".pdf")
            result = convert_xlsx_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "CSV to PDF":
            output_file = tmp_file_path.replace(".csv", ".pdf")
            result = convert_csv_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "JSON to PDF":
            output_file = tmp_file_path.replace(".json", ".pdf")
            result = convert_json_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "XML to PDF":
            output_file = tmp_file_path.replace(".xml", ".pdf")
            result = convert_xml_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "TXT to CSV":
            output_file = tmp_file_path.replace(".txt", ".csv")
            result = convert_txt_to_csv(tmp_file_path, output_file)
        elif conversion_type == "Compress Folder":
            output_file = tmp_file_path.replace(".zip", "")
            result = compress_folder(tmp_file_path, output_file)
        elif conversion_type == "Extract ZIP":
            output_folder = tmp_file_path.replace(".zip", "")
            result = extract_zip(tmp_file_path, output_folder)
        else:
            result = "Invalid conversion type selected."

        if "Error" in result:
            st.error(result)
        else:
            st.success(f"Conversion successful! Download the file below.")
            with open(output_file, "rb") as f:
                st.download_button("Download Converted File", f, file_name=os.path.basename(output_file))

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

# Optional: Add instructions or descriptions
st.write("This platform allows you to convert a variety of file formats. Upload your files and choose the desired conversion type.")
