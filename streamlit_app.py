import streamlit as st
import os
from tempfile import NamedTemporaryFile
from fpdf import FPDF
import markdown
from PIL import Image
import json
import xml.etree.ElementTree as ET
import pyqrcode
import barcode
from barcode.writer import ImageWriter
import pytz
from datetime import datetime
from io import BytesIO
import pandas as pd
import pyttsx3
import speech_recognition as sr
from moviepy.editor import VideoFileClip
import ebooklib
from ebooklib import epub
import shutil
import zipfile
from pptx import Presentation

# Streamlit App UI Setup
st.title("Advanced File Conversion Platform with AI Assistant")
st.write("Convert various file formats and get assistance from the AI assistant.")

# File upload section
uploaded_file = st.file_uploader("Upload your file", type=["dwg", "rvt", "ai", "fdr", "pptx", "txt", "zip", "jpg", "jpeg", "png", "md", "html", "epub", "json", "xml", "mp3", "mp4"])

# Conversion options
conversion_type = st.selectbox(
    "Select the conversion format",
    [
        "DWG to PDF", "RVT to DWG", "AI to PDF", "FDR to PDF", "PPT to PDF", "TXT to PDF", 
        "MD to PDF", "HTML to PDF", "EPUB to PDF", "JSON to PDF", "XML to PDF", "Extract ZIP", 
        "Compress Folder", "Image to PDF", "Audio to Text", "Text to Speech", "Video to Audio", 
        "QR Code Generator", "Barcode Generator", "HTML to Text", "CSV to Excel", "Excel to CSV",
        "Markdown to Text", "Text to Markdown", "Text File Cleaner"
    ]
)

# File conversion functions
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

def text_to_speech(input_text, output_audio):
    try:
        engine = pyttsx3.init()
        engine.save_to_file(input_text, output_audio)
        engine.runAndWait()
        return output_audio
    except Exception as e:
        return f"Error in text to speech conversion: {str(e)}"

def convert_video_to_audio(input_file, output_audio):
    try:
        video_clip = VideoFileClip(input_file)
        video_clip.audio.write_audiofile(output_audio)
        return output_audio
    except Exception as e:
        return f"Error in video to audio conversion: {str(e)}"

def extract_zip(input_file, output_folder):
    try:
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(output_folder)
        return f"Extracted to {output_folder}"
    except Exception as e:
        return f"Error in zip extraction: {str(e)}"

def compress_folder(input_folder, output_file):
    try:
        shutil.make_archive(output_file, 'zip', input_folder)
        return f"Folder compressed to {output_file}.zip"
    except Exception as e:
        return f"Error in folder compression: {str(e)}"

# Handle file upload and conversion process
if uploaded_file:
    # Process uploaded file
    with NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name

    result = "No conversion performed"  # Default value for result
    output_file = None  # Initialize output_file as None

    try:
        if conversion_type == "TXT to PDF":
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
        elif conversion_type == "JSON to PDF":
            output_file = tmp_file_path.replace(".json", ".pdf")
            result = convert_json_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "XML to PDF":
            output_file = tmp_file_path.replace(".xml", ".pdf")
            result = convert_xml_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "QR Code Generator":
            output_file = tmp_file_path.replace(".txt", ".png")
            result = generate_qr_code(tmp_file_path, output_file)
        elif conversion_type == "Barcode Generator":
            output_file = tmp_file_path.replace(".txt", ".png")
            result = generate_barcode(tmp_file_path, output_file)
        elif conversion_type == "Image to PDF":
            output_file = tmp_file_path.replace(".jpg", ".pdf")
            result = convert_image_to_pdf(tmp_file_path, output_file)
        elif conversion_type == "Audio to Text":
            result = convert_audio_to_text(tmp_file_path)
        elif conversion_type == "Text to Speech":
            output_file = tmp_file_path.replace(".txt", ".mp3")
            result = text_to_speech(tmp_file_path, output_file)
        elif conversion_type == "Video to Audio":
            output_file = tmp_file_path.replace(".mp4", ".mp3")
            result = convert_video_to_audio(tmp_file_path, output_file)
        elif conversion_type == "Extract ZIP":
            output_folder = tmp_file_path.replace(".zip", "")
            result = extract_zip(tmp_file_path, output_folder)
        elif conversion_type == "Compress Folder":
            output_file = tmp_file_path.replace(".zip", "_compressed")
            result = compress_folder(tmp_file_path, output_file)

    except Exception as e:
        result = f"Error: {str(e)}"

    # Provide feedback to the user
    if output_file:
        st.download_button(
            label="Download the converted file",
            data=open(output_file, 'rb').read(),
            file_name=output_file,
            mime="application/octet-stream"
        )
    st.write(result)

