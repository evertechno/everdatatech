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
import io

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

# New Advanced Functions
def csv_to_excel(input_file, output_file):
    try:
        df = pd.read_csv(input_file)
        df.to_excel(output_file, index=False)
        return output_file
    except Exception as e:
        return f"Error in CSV to Excel conversion: {str(e)}"

def excel_to_csv(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        df.to_csv(output_file, index=False)
        return output_file
    except Exception as e:
        return f"Error in Excel to CSV conversion: {str(e)}"

def markdown_to_text(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            markdown_content = file.read()
        text_content = markdown.markdown(markdown_content)
        with open(output_file, 'w') as file:
            file.write(text_content)
        return output_file
    except Exception as e:
        return f"Error in Markdown to Text conversion: {str(e)}"

def text_to_markdown(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            text_content = file.read()
        markdown_content = markdown.markdown(text_content)
        with open(output_file, 'w') as file:
            file.write(markdown_content)
        return output_file
    except Exception as e:
        return f"Error in Text to Markdown conversion: {str(e)}"

def text_file_cleaner(input_file, output_file):
    try:
        with open(input_file, 'r') as file:
            content = file.read()
        cleaned_content = content.strip().replace("\n", " ").replace("\r", "")
        with open(output_file, 'w') as file:
            file.write(cleaned_content)
        return output_file
    except Exception as e:
        return f"Error in Text File Cleaner: {str(e)}"

# Handle file upload and processing
if uploaded_file is not None:
    # Handle different conversion options
    if conversion_type == "PPT to PDF":
        output_file = convert_ppt_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "TXT to PDF":
        output_file = convert_txt_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "MD to PDF":
        output_file = convert_md_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "HTML to PDF":
        output_file = convert_html_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "EPUB to PDF":
        output_file = convert_epub_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "JSON to PDF":
        output_file = convert_json_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "XML to PDF":
        output_file = convert_xml_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "Image to PDF":
        output_file = convert_image_to_pdf(uploaded_file, "output.pdf")
    elif conversion_type == "QR Code Generator":
        output_file = generate_qr_code("sample text", "output_qr.png")
    elif conversion_type == "Barcode Generator":
        output_file = generate_barcode("123456789012", "output_barcode.png")
    elif conversion_type == "Audio to Text":
        output_file = convert_audio_to_text(uploaded_file)
    elif conversion_type == "Text to Speech":
        output_file = text_to_speech("Hello world!", "output_audio.mp3")
    elif conversion_type == "Video to Audio":
        output_file = convert_video_to_audio(uploaded_file, "output_audio.mp3")
    elif conversion_type == "Text File Cleaner":
        output_file = text_file_cleaner(uploaded_file, "cleaned_output.txt")
    else:
        output_file = "Unsupported conversion type."

    st.download_button("Download Output", output_file)

