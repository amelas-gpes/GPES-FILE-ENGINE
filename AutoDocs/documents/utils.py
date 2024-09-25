# Import necessary libraries for PDF generation, data handling, and user interaction
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, legal
from reportlab.lib.pagesizes import A1, A2, A3, A4, A5, A6, landscape
from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, Flowable
)
from io import BytesIO
import PyPDF2
from reportlab.pdfgen import canvas
import random  # For generating random data
import pandas as pd  # For reading data from Excel files
from datetime import datetime  # For current date and time
import os  # For file path operations
import re  # For regular expressions
import shutil
import fitz

def generate_bank_name():
    # List of banks to choose from
    banks = [
        "JPMorgan Chase", "Bank of America", "Wells Fargo", "Citibank",
        "Goldman Sachs", "Morgan Stanley", "PNC Financial Services",
        "U.S. Bancorp", "TD Bank", "Capital One", "HSBC", "Barclays",
        "Deutsche Bank", "Credit Suisse", "BNP Paribas", "UBS",
        "Santander", "Royal Bank of Canada", "Scotiabank", "ING Group"
    ]

    return random.choice(banks)

def generate_account_number():
    return random.randrange(1_000_000_000, 10_000_000_000)
    
def generate_ABA():
    return random.randrange(100_000_000, 1_000_000_000)

def generate_SWIFT():
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    
    return ''.join(random.choice(letters) for _ in range(5)) + \
        ''.join(str(random.randrange(10)) for _ in range(2))

# Function to create a logo image while maintaining aspect ratio
def create_logo(max_width, max_height, image_path):
    """
    Creates an Image object for the logo that fits within the specified
    dimensions while maintaining its aspect ratio.
    """
    # Create an Image object using the provided image path
    logo = Image(image_path)
    logo.hAlign = 'LEFT'  # Align the logo to the left

    # Get the original dimensions of the image
    image_width, image_height = logo.wrap(0, 0)

    # Calculate the scaling factor to fit the image within the max dimensions
    scale = min(max_width / image_width, max_height / image_height, 1.0)

    # Apply the scaling factor
    logo.drawWidth = image_width * scale
    logo.drawHeight = image_height * scale

    return logo

# Function to format a phone number in the (xxx) xxx-xxxx format
def format_phone_number(num1, num2, num3):
    """
    Formats a phone number into the (xxx) xxx-xxxx format.
    """
    formatted_number = "({}) {}-{}".format(num1, num2, num3)
    return formatted_number

# Function to create a horizontal divider line
class DividerLine(Flowable):
    """
    A custom Flowable to draw a horizontal divider line.
    """
    def __init__(self, width, thickness=0.5, color=colors.lightgrey):
        super().__init__()
        self.width = width
        self.thickness = thickness
        self.color = color

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# Function to sanitize filenames by removing invalid characters
def sanitize_filename(filename):
    """
    Removes invalid characters from filenames using regex.
    """
    # Define a regex pattern for invalid characters
    invalid_chars = r'[<>:"/\\|?*]'
    # Replace invalid characters with an underscore
    return re.sub(invalid_chars, '_', filename)

def add_multiple_texts_to_existing_pdf(input_pdf, output_pdf, texts_with_positions):
    existing_pdf = PyPDF2.PdfReader(input_pdf)

    # Step 1: Create a PDF with all the texts you want to add
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    
    # Loop through the list of texts and positions
    for text, position in texts_with_positions:
        x, y = position
        can.setFont("Helvetica", 8)
        can.setFillColorRGB(0, 0, 0)
        can.drawString(x, y, text)  # Add each text at the specified (x, y) position
    
    can.save()
    
    # Move to the beginning of the BytesIO buffer
    packet.seek(0)

    # Step 3: Read the new PDF (with the added text)
    new_pdf = PyPDF2.PdfReader(packet)

    # Step 4: Merge the new text with the existing PDF
    output = PyPDF2.PdfWriter()

    # Iterate through each page of the existing PDF and merge the text
    for i in range(len(existing_pdf.pages)):
        page = existing_pdf.pages[i]
        if i == 0:  # Assuming you want to add text to the first page
            page.merge_page(new_pdf.pages[0])  # Merge the text PDF onto the first page
        output.add_page(page)

    # Step 5: Save the result to a new PDF file
    with open(output_pdf, "wb") as output_file:
        output.write(output_file)

def split_by_length(strings, max_length):
    chunks = []
    current_chunk = ""
    current_length = 0

    for string in strings:
        if current_length + len(string) <= max_length:
            current_chunk += string + " "
            current_length += len(string)
        else:
            chunks.append(current_chunk)
            current_chunk = string + " "
            current_length = len(string)

    if current_chunk:
        chunks.append(current_chunk)

    return chunks