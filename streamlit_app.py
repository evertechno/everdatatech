import streamlit as st
import os
from tempfile import NamedTemporaryFile
from PyPDF2 import PdfMerger
from pptx import Presentation
from fpdf import FPDF
import markdown
from ebooklib import epub
from docx import Document
import zipfile
import shutil
from PIL import Image
import json
import xml.etree.ElementTree as ET
import cairosvg
import csv
import pandas as pd
import wave
import numpy as np
import pyttsx3
from datetime import datetime
import pyqrcode
from io import BytesIO
import barcode
from barcode.writer import ImageWriter
import pytz
from lxml import etree
import subprocess

# Remove external dependencies for things like AutoCAD, Google Gemini, and other unsupported modules

# Conversion Functions (Updated)

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

def extract_zip(input_file, output_folder):
    try:
        with zipfile.ZipFile(input_file, 'r') as zip_ref:
            zip_ref.extractall(output_folder)
        return f"Extracted zip to {output_folder}"
    except Exception as e:
        return f"Error in extracting ZIP: {str(e)}"

def compress_folder(input_folder, output_file):
    try:
        shutil.make_archive(output_file, 'zip', input_folder)
        return output_file + ".zip"
    except Exception as e:
        return f"Error in compressing folder: {str(e)}"

def convert_image_to_pdf(input_file, output_file):
    try:
        image = Image.open(input_file)
        image.convert('RGB').save(output_file, "PDF")
        return output_file
    except Exception as e:
        return f"Error in Image to PDF conversion: {str(e)}"

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
        return f"Error converting HTML to text: {str(e)}"

# Add more functions as needed based on the required conversions
