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
from tkinter import Tk, filedialog  # For file and directory selection dialogs
import re  # For regular expressions
import shutil
import fitz

# List of banks to choose from
banks = [
    "JPMorgan Chase", "Bank of America", "Wells Fargo", "Citibank",
    "Goldman Sachs", "Morgan Stanley", "PNC Financial Services",
    "U.S. Bancorp", "TD Bank", "Capital One", "HSBC", "Barclays",
    "Deutsche Bank", "Credit Suisse", "BNP Paribas", "UBS",
    "Santander", "Royal Bank of Canada", "Scotiabank", "ING Group"
]

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

# Main function to create the capital call PDF
def create_capital_call_pdf(filename, investing_entity_name, legal_name, image_path):
    """
    Generates a capital call PDF document with the given filename and content.

    Parameters:
    - filename: Name of the output PDF file.
    - investing_entity_name: Name of the investing entity (fund).
    - legal_name: Legal name of the investor.
    - image_path: Path to the logo image file.
    """
    # Create a document template with specified margins and page size
    doc = SimpleDocTemplate(
        filename, pagesize=letter,
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch
    )

    story = []  # List to hold the flowable elements for the PDF

    # Define styles for the document
    styles = getSampleStyleSheet()

    # Update default style with built-in font
    styles['Normal'].fontName = 'Helvetica'
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 12

    # Custom styles for headings and body text
    styles.add(ParagraphStyle(
        name='CustomHeading1',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=10,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomHeading2',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceBefore=14,
        spaceAfter=4,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomBody',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name='CustomEmphasis',
        parent=styles['CustomBody'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#000000"),
    ))
    # Style for table cells
    table_cell_style = ParagraphStyle(
        name='TableCell',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
    )
    table_cell_bold_style = ParagraphStyle(
        name='TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold',
    )

    # Add Logo to the document with aspect ratio maintained
    logo = create_logo(80, 40, image_path)  # Adjust max dimensions as needed
    story.append(logo)
    story.append(Spacer(1, 10))  # Add space after the logo

    # Add Title
    story.append(Paragraph("Capital Call Notice", styles['CustomHeading1']))

    # Use the current date for the notice
    current_date = datetime.now().strftime("%B %d, %Y")  # Format: Month Day, Year

    # Create header information with date, recipient, and sender details
    header_info = f"""
    <b>Date:</b> {current_date}<br/>
    <b>To:</b> {legal_name}<br/>
    <b>From:</b> {investing_entity_name}
    """
    story.append(Paragraph(header_info.strip(), styles['CustomBody']))

    # Divider line
    story.append(Spacer(1, 6))
    story.append(DividerLine(doc.width))
    story.append(Spacer(1, 6))

    # Company Description
    company_description = f"""
    {investing_entity_name} is a private equity investment fund focused on identifying and nurturing high-potential growth companies across various sectors. With a proven track record of strategic investments and value creation, our fund aims to generate superior returns for our limited partners while fostering innovation and sustainable growth in our portfolio companies.
    """
    story.append(Paragraph(company_description.strip(), styles['CustomBody']))

    # Capital Call Details Section
    story.append(Spacer(1, 4))
    story.append(Paragraph("Capital Call Details", styles['CustomHeading2']))

    # Generate financial data for the capital call
    total_commitment = random.randrange(800_000, 1_000_000)
    prior_capital_contributions = random.randrange(600_000, total_commitment)
    unfunded_commitment = total_commitment - prior_capital_contributions
    current_capital_call = random.randrange(0, unfunded_commitment)
    remaining_unfunded_commitment = unfunded_commitment - current_capital_call

    # Prepare data for the financial table with emphasis on key numbers
    data = [
        [
            Paragraph("Description", table_cell_bold_style),
            Paragraph("Amount", table_cell_bold_style)
        ],
        [
            Paragraph("Total Commitment", table_cell_style),
            Paragraph("${:,}".format(total_commitment), table_cell_bold_style)
        ],
        [
            Paragraph("Prior Capital Contributions", table_cell_style),
            Paragraph("${:,}".format(prior_capital_contributions), table_cell_style)
        ],
        [
            Paragraph("Unfunded Commitment", table_cell_style),
            Paragraph("${:,}".format(unfunded_commitment), table_cell_bold_style)
        ],
        [
            Paragraph("Current Capital Call", table_cell_style),
            Paragraph("${:,}".format(current_capital_call), table_cell_bold_style)
        ],
        [
            Paragraph("Remaining Unfunded Commitment", table_cell_style),
            Paragraph("${:,}".format(remaining_unfunded_commitment), table_cell_style)
        ],
    ]

    # Create and style the financial table
    table = Table(data, colWidths=[3.5 * inch, 2 * inch], hAlign='LEFT', repeatRows=1)

    # Table style with alternate row shading and thin lines
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#333333")),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Align all cells to the left
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
        ('FONTSIZE', (0, 0), (-1, -1), 10),  # Font size for all cells
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#fafafa")]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#dddddd")),
    ]))
    story.append(table)

    # Purpose of Capital Call Section
    story.append(Spacer(1, 10))
    story.append(Paragraph("Purpose of Capital Call", styles['CustomHeading2']))
    purpose = """
    • New Portfolio Company Investment: <b>${:,}</b><br/>
    • Follow-on Investment: <b>${:,}</b><br/>
    • Fund Expenses: <b>${:,}</b>
    """.format(
        random.randrange(600_000, 1_000_000),
        random.randrange(600_000, 1_000_000),
        random.randrange(600_000, 1_000_000)
    )
    story.append(Paragraph(purpose.strip(), styles['CustomBody']))

    # Payment Instructions Section
    story.append(Spacer(1, 10))
    story.append(Paragraph("Payment Instructions", styles['CustomHeading2']))

    # Generate random account details
    account_number = generate_account_number()  
    ABA = generate_ABA() 
    SWIFT = generate_SWIFT()



    # Emphasize account details
    payment = f"""
    <b>Bank:</b> {random.choice(banks)}<br/>
    <b>Account:</b> {investing_entity_name}<br/>
    <b>Account Number:</b> <b>{account_number}</b><br/>
    <b>ABA:</b> <b>{ABA}</b><br/>
    <b>SWIFT:</b> {SWIFT}
    """
    story.append(Paragraph(payment.strip(), styles['CustomBody']))

    # Due Date Section with emphasis
    due_date = (datetime.now() + pd.DateOffset(days=7)).strftime("%B %d, %Y")
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        f"<b>Due Date:</b> <b>{due_date}</b>",
        styles['CustomBody']
    ))

    # Divider line before footer
    story.append(Spacer(1, 14))
    story.append(DividerLine(doc.width))
    story.append(Spacer(1, 6))

    # Footer with contact information
    legal_names = [
        "Alexander Johnson", "Victoria Smith", "Benjamin Clarke",
        "Elizabeth Baker", "Christopher Miller", "Sophia Davis",
        "Matthew Wilson", "Isabella Moore", "William Taylor",
        "Charlotte Anderson", "Nicholas Thomas", "Catherine Brown",
        "Jonathan Harris", "Margaret Martin", "Daniel Thompson",
        "Rebecca Lee", "Michael Scott", "Alexandra White",
        "Robert Lewis", "Katherine Roberts"
    ]
    contact_role = "IR Manager"

    # Select a random contact name
    contact_name = random.choice(legal_names)

    # Format the contact information
    contact = f"""
    <b>Contact:</b> {contact_name}, {contact_role}<br/>
    {contact_name.replace(' ', '.').lower()}@{investing_entity_name.lower().replace(' ', '')}.com | {format_phone_number(random.randrange(100, 1000), random.randrange(100, 1000), random.randrange(1000, 10000))}
    """
    story.append(Paragraph(contact.strip(), styles['CustomBody']))

    # Build the PDF document
    doc.build(story)


def add_multiple_texts_to_existing_pdf(input_pdf, output_pdf, texts_with_positions):
    existing_pdf = PyPDF2.PdfReader(input_pdf)

    # Step 1: Create a PDF with all the texts you want to add
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    
    # Loop through the list of texts and positions
    for text, position in texts_with_positions:
        x, y = position
        can.setFont("Helvetica-Bold", 20)
        can.setFillColorRGB(1, 1, 1)
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


# Main function to create the quarterly update PDF
def create_quarterly_update_pdf(filename, investing_entity_name, legal_name, image_path):
    """
    Generates a capital call PDF document with the given filename and content.

    Parameters:
    - filename: Name of the output PDF file.
    - investing_entity_name: Name of the investing entity (fund).
    - legal_name: Legal name of the investor.
    - image_path: Path to the logo image file.
    """
    # Create a document template with specified margins and page size
    doc = SimpleDocTemplate(
        filename, pagesize=letter,
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch
    )

    story = []  # List to hold the flowable elements for the PDF

    # Define styles for the document
    styles = getSampleStyleSheet()

    # Update default style with built-in font
    styles['Normal'].fontName = 'Helvetica'
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 12

    # Custom styles for headings and body text
    styles.add(ParagraphStyle(
        name='CustomHeading1',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=10,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomHeading2',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceBefore=14,
        spaceAfter=4,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomBody',
        fontName='Helvetica',
        fontSize=12,
        leading=12,
        textColor=colors.HexColor("#515154"),
        spaceAfter=12,
    ))
    styles.add(ParagraphStyle(
        name='CustomEmphasis',
        parent=styles['CustomBody'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#000000"),
    ))
    # Style for table cells
    table_cell_style = ParagraphStyle(
        name='TableCell',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
    )
    table_cell_bold_style = ParagraphStyle(
        name='TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold',
    )

    # Add Logo to the document with aspect ratio maintained
    logo = create_logo(80, 40, image_path)  # Adjust max dimensions as needed
    story.append(logo)
    story.append(Spacer(1, 10))  # Add space after the logo

    # Add Title
    story.append(Paragraph("July 31, 2024", styles['CustomHeading2']))
    story.append(Paragraph(f"{investing_entity_name} Partners - Quarterly Update Q3 2024", styles['CustomHeading2']))
    story.append(Paragraph(f"{investing_entity_name} Limited Partners,", styles['CustomHeading2']))

    # Use the current date for the notice
    current_date = datetime.now().strftime("%B %d, %Y")  # Format: Month Day, Year

    # Divider line
    story.append(Spacer(1, 6))
    story.append(DividerLine(doc.width))
    story.append(Spacer(1, 6))

    # Company Description
    company_description = f"""
    
    In the third quarter of this year, our company experienced significant growth across key performance areas, building on the strong momentum we’ve established over the past several quarters. Revenue increased by 12%, largely driven by heightened demand for our flagship product line and our ongoing strategic expansion into new geographic markets. Our international division, in particular, saw robust performance, with a 15% increase in global sales, especially from the Asia-Pacific and European regions. This growth was bolstered by our decision to localize products for these markets, a strategy that has resonated well with customers. Additionally, the company achieved a 5% reduction in overhead costs through improved operational efficiencies. By streamlining processes, leveraging automation tools, and optimizing our supply chain, we’ve made significant strides in controlling costs while maintaining quality and service excellence.
    """
    story.append(Paragraph(company_description.strip(), styles['CustomBody']))

    paragraph2 = f"""
    From an innovation standpoint, this quarter marked the successful launch of two new product lines, which have already begun gaining traction in the marketplace. These products were developed with a strong emphasis on sustainability and environmental impact, which aligns with our company’s long-term commitment to corporate responsibility. Early customer feedback has been overwhelmingly positive, with many praising the eco-friendly features as a standout factor. We believe that these new products will not only meet the growing consumer demand for sustainable options but will also solidify our position as a leader in responsible manufacturing. Moreover, we’ve increased our investment in research and development by 20%, further reinforcing our commitment to innovation. This additional funding has been allocated toward advancing new technologies and enhancing our existing product offerings to ensure we remain competitive in a rapidly evolving industry landscape.
    """
    story.append(Paragraph(paragraph2.strip(), styles['CustomBody']))

    paragraph3 = f"""
    As we look ahead to the final quarter of the fiscal year, our primary focus will be on sustaining this momentum and driving growth through several key initiatives. One of our top priorities will be advancing our digital transformation strategy. This includes launching new digital tools aimed at improving customer experience, streamlining sales processes, and optimizing supply chain management. In addition to these external-facing initiatives, we are also deeply committed to fostering a more inclusive and diverse workplace. In the upcoming quarter, we will be rolling out several new internal programs designed to promote diversity, equity, and inclusion within our workforce. These programs are aimed at supporting underrepresented groups, enhancing leadership diversity, and fostering a more collaborative and equitable company culture. Together, these efforts will position us for continued success as we close out the year on a strong note.
    """
    story.append(Paragraph(paragraph3.strip(), styles['CustomBody']))

    # Build the PDF document
    doc.build(story)

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

"""
def create_gp_report_pdf(filename, text_to_add, output_pdf_path, all_fund_codes):
    # Open the existing PDF
    doc = fitz.open(filename)
    page = doc.load_page(0)  # Select the first page

    logo_blocker = fitz.Rect(5, 20, 200, 60)    
    name_blocker = fitz.Rect(10, 110, 70, 130)
    footer_blocker = fitz.Rect(5, 570, 600, 600)
    date_blocker = fitz.Rect(90, 95, 170, 105)

    
    page.draw_rect(name_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(logo_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(footer_blocker, color=(1,1,1), fill=(1,1,1))
    page.draw_rect(date_blocker, color=(1,1,1), fill=(1,1,1))

    # logo
    page.insert_image(logo_blocker, filename=text_to_add["logo"])
    # legal name
    page.insert_text((15, 125), text_to_add["legal_name"], fontsize=12, color=(0,0,0))
    # footer
    page.insert_text((15, 590), text_to_add["footer"], fontsize=12, color=(0,0,0))
    # date
    page.insert_text((90, 102), text_to_add["date"], fontsize=8, color=(0,0,0))

    all_fund_codes = list(all_fund_codes)

    fund_x_pos = 230
    fund = random.choice(all_fund_codes)
    fund_y_pos = 141
    page.insert_text((fund_x_pos - (len(fund) * 5)+ 5, fund_y_pos), fund, fontsize=8, color=(1,0,0))

    fund_x_pos += 70
    fund = random.choice(all_fund_codes)
    fund_y_pos = 141
    page.insert_text((fund_x_pos - (len(fund) * 5)+ 5, fund_y_pos), fund, fontsize=8, color=(1,0,0))


    fund1 = ["-","-","-", "", "", "-","-","-", "", "", "-","-","-","-","-","-","-","2,049","2,049","","","-","-","-","-","2,049","2,049","2,049"]
    fund1_num1 = random.randrange(1000, 9999)
    fund1[17] = "{:,}".format(fund1_num1)
    fund1[18] = "{:,}".format(fund1_num1)
    fund1[25] = "{:,}".format(fund1_num1)
    fund1[26] = "{:,}".format(fund1_num1)
    fund1[27] = "{:,}".format(fund1_num1)

    fund2 = ["-","-","-", "", "", "-","-","-", "", "", "-","-","-","-","-","-","-","-","-","","","-","-","-","-","-","-","-"]
    total = ["-","-","-", "", "", "-","-","-", "", "", "-","-","-","-","-","-","-","-","-","","","-","-","-","-","-","-","-"]

    fund2_num1 = random.randrange(10,99)
    fund2[15] = "({fund2_num1})"
    fund2[16] = fund2_num1
    fund2[22] = fund2_num1
    fund2[23] = fund2_num1

    total[15] = fund2_num1
    total[16] = fund2_num1

    fund2_num2 = random.randrange(1000, 9999)
    fund2[17] = "{:,}".format(fund2_num2)
    fund2[18] = "{:,}".format(fund2_num2)

    fund2[26] = "{:,}".format(fund2_num2)
    fund2[27] = "{:,}".format(fund2_num2)
    fund2[28] = "{:,}".format(fund1_num1 + fund2_num2)

    total[17] = fund2_num1
    total[18] = fund2_num1
    total[19] = fund1_num1 + fund2_num2 
    total[20] = fund1_num1 + fund2_num2 

    total[22] = fund2_num1
    total[23] = fund2_num1

    total[25] = fund1_num1 + fund2_num2 
    total[26] = fund1_num1 + fund2_num2 
    total[27] = fund1_num1 + fund2_num1 + fund2_num2
    
    x_pos = 230
    y_pos = 152
    for num in leof:
        page.insert_text((x_pos - (len(num) * 3)+3, y_pos), num, fontsize=8, color=(1,0,0))
        y_pos += 11.1

    x_pos += 63
    y_pos = 152
    for num in lsc_2:
        page.insert_text((x_pos - (len(num) * 3)+3, y_pos), num, fontsize=8, color=(1,0,0))
        y_pos += 11.1

    x_pos += 63
    y_pos = 152
    for num in total:
        page.insert_text((x_pos - (len(num) * 3)+3, y_pos), num, fontsize=8, color=(1,0,0))
        y_pos += 11.1

    # Save the modified PDF
    doc.save(output_pdf_path)
    doc.close()
"""
from reportlab.pdfbase import pdfmetrics


def calculate_col_widths(text, font_name='Helvetica', font_size=10):
    text_width = pdfmetrics.stringWidth(str(text), font_name, font_size)
    print(text_width)
    return text_width - 10  # Add a little padding


def create_gp_report_pdf(filename, investing_entity, legal_name, image_path, fund_names, footer):
    """
    Generates a gp report PDF document with the given filename and content.

    Parameters:
    - filename: Name of the output PDF file.
    - investing_entity_name: Name of the investing entity (fund).
    - legal_name: Legal name of the investor.
    - image_path: Path to the logo image file.
    """
    # Create a document template with specified margins and page size
    doc = SimpleDocTemplate(
        filename, pagesize=landscape(A3),
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch
    )

    story = []  # List to hold the flowable elements for the PDF

    # Define styles for the document
    styles = getSampleStyleSheet()

    # Update default style with built-in font
    styles['Normal'].fontName = 'Helvetica'
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 12

    # Custom styles for headings and body text
    styles.add(ParagraphStyle(
        name='CustomHeading1',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=10,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomHeading2',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceBefore=14,
        spaceAfter=4,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomBody',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name='CustomEmphasis',
        parent=styles['CustomBody'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#000000"),
    ))
    # Style for table cells
    table_cell_style = ParagraphStyle(
        name='TableCell',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
    )
    table_cell_bold_style = ParagraphStyle(
        name='TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold',
    )

    # Add Logo to the document with aspect ratio maintained
    logo = create_logo(80, 40, image_path)  # Adjust max dimensions as needed
    story.append(logo)
    story.append(Spacer(1, 10))  # Add space after the logo

    # Use the current date for the notice
    current_date = datetime.now().strftime("%B %d, %Y")  # Format: Month Day, Year

    # Create header information with date, recipient, and sender details
    header_info = f"""
    <b>General Partners of the Fund(s)</b><br/>
    Unaudited Capital Account Statement<br/>
    As of and life to date {current_date}<br/>
    <br/>
    <b>{legal_name}</b>
    """
    story.append(Paragraph(header_info.strip(), styles['CustomBody']))

    # Prepare data for the financial table with emphasis on key numbers
    data = [
    ]

    #29
    row_titles = ["<b>Commitment Summary</b>", "Investor commitment", "Gross contributions", "Remaining commitment", "", 
    "<b>Waiver Analysis</b>", "Deemed contributions", "Realized special profit", "Remaining special profit", "", 
    "<b>Life to Date Book Capital Account</b>", "Capital contributions", "Net investment distributions", "Net income / (loss) & syndication costs",
    "Net realized investment gain / (loss)", "Net unrealized investment gain / (loss)", "Net carried interest distributions", "Net realized carried interest",
    "Net unrealized carried interest", "<b>Ending Capital Account Balance</b>", "",
    "<b>Summary</b>", "Investment distributions", "Carried interest distributions", "<b>Total distributions</b>", "Investment Current Value", "Carried interest Current Value", 
    "<b>Total Current Value</b>", "<b>Total Value (Capital + Distributions)</b>"]

    currency_col = ["", "$", "", "$", "", 
    "", "$", "", "$", "", 
    "", "$", "", "",
    "", "", "", "",
    "", "$", "",
    "", "$", "", "$", "$", "", 
    "$", "$"]

    fund1_num1 = random.randrange(1000, 9999)


    fund1_name = random.choice(fund_names)
    while (fund1_name == investing_entity):
        fund1_name = random.choice(fund_names)

    fund2_name = fund1_name

    while (fund2_name == fund1_name):
        fund2_name = random.choice(fund_names)


    fund1 = ["<b>" + fund1_name + "</b>", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", "-", "-",
    f"{fund1_num1}", f"{fund1_num1}", "",
    "", "-", "-", "-", "-", f"{fund1_num1}", 
    f"{fund1_num1}", f"{fund1_num1}"]

    fund2_num1 = random.randrange(10, 999)
    fund2_num2 = random.randrange(1000, 9999)

    fund2 = ["<b>" + fund2_name + "</b>", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", f"({fund2_num1})", f"{fund2_num1}",
    f"{fund2_num2}", f"{fund2_num2}", "",
    "", "-", f"{fund2_num1}", f"{fund2_num1}", "-", f"{fund2_num2}", 
    f"{fund2_num2}", f"{fund2_num2}"]

    fund2[28] = str(int(fund2[27]) + int(fund2[24]))

    fund_info = []
    fund_info.append(
        fund1
    )
    fund_info.append(
        fund2
    )

    total = ["<b>TOTAL</b>", "-", "-", "-", "", 
    "", "-", "-", "-", "", 
    "", "-", "-", "-",
    "-", "-", "(63)", "63",
    "3419", "3419", "",
    "", "-", "63", "63", "-", "3419", 
    "3419", "3482"]

    for i in range(len(fund1)):
        if (fund1[i].isdigit() and fund2[i].isdigit()):
            total[i] = str(int(fund1[i]) + int(fund2[i]))
        elif (fund1[i].isdigit()):
            total[i] = fund1[i]
        elif (fund2[i].isdigit()):
            total[i] = fund2[i]
        elif (fund1[i][1:-1].isdigit()):
            total[i] = fund1[i]
        elif (fund2[i][1:-1].isdigit()):
            total[i] = fund2[i]


    for index in range(len(row_titles)):
        row = []

        row.append(Paragraph(row_titles[index]))
        row.append(Paragraph(currency_col[index]))

        for j in range(len(fund_info)):
            if index >= len(fund_info[j]):
                row.append(Paragraph(""))
                continue
            row.append(Paragraph(fund_info[j][index]))
        
        row.append(Paragraph(total[index]))

        data.append(row)

    # Create and style the financial table
    font_name = "Helvetica"
    font_size = 8

    col_widths = [3*inch, .5 * inch]
    for i in range(len(fund_info)):
        col_widths.append(3 * inch)

    

    table = Table(data, colWidths=col_widths, hAlign='LEFT', repeatRows=1)

    style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Set font size to 8 for all cells
        ('BOX', (0,0), (-1,-1), 2, colors.black),
        ('GRID', (0,1), (-1,-1), 1, colors.black)
    ])
    table.setStyle(style)

    story.append(table)

    story.append(Spacer(1, 10))

    story.append(Paragraph(footer, styles['CustomHeading2']))
    
    # Build the PDF document
    doc.build(story)

"""
def create_wire_instruction_pdf(filename, text_to_add, output_pdf_path):
    doc = fitz.open(filename)
    page = doc.load_page(0)

    logo_blocker = fitz.Rect(15, 20, 250, 70)    
    date_blocker = fitz.Rect(112, 195, 250, 210)
    name_blocker = fitz.Rect(90, 220, 170, 240)
    fund_name_blocker = fitz.Rect(407, 170, 600, 183)

    page.draw_rect(name_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(logo_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(date_blocker, color=(1,1,1), fill=(1,1,1))
    page.draw_rect(fund_name_blocker, color=(1,1,1), fill=(1,1,1))

    # logo
    page.insert_image(logo_blocker, filename=text_to_add["logo"])
    # legal name
    page.insert_text((90, 232), text_to_add["legal_name"], fontsize=12, color=(0,0,0))
    # fund name
    page.insert_text((407, 180), text_to_add["fund_name"], fontsize=10, color=(0,0,0))
    # date
    page.insert_text((112, 207), text_to_add["date"], fontsize=10, color=(0,0,0))

    
    #Bank name
    page.insert_text((205, 305), random.choice(banks), fontsize=12, color=(0,0,0))
    #ABA/SWIFT
    page.insert_text((205, 320), generate_SWIFT(), fontsize=12, color=(0,0,0))
    #Beneficiary Name
    page.insert_text((205, 378), text_to_add["legal_name"], fontsize=12, color=(0,0,0))
    #Account/Iban
    page.insert_text((205, 393), str(generate_account_number()), fontsize=12, color=(0,0,0))
    #address
    page.insert_text((205, 440), text_to_add["address_1"], fontsize=12, color=(0,0,0))


    doc.save(output_pdf_path)
    doc.close()
"""

def create_wire_instruction_pdf(text_to_add, output_pdf_path):
    """
    Generates a gp report PDF document with the given filename and content.

    Parameters:
    - filename: Name of the output PDF file.
    - investing_entity_name: Name of the investing entity (fund).
    - legal_name: Legal name of the investor.
    - image_path: Path to the logo image file.
    """
    # Create a document template with specified margins and page size
    doc = SimpleDocTemplate(
        output_pdf_path, pagesize=letter,
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.75 * inch, bottomMargin=0.75 * inch
    )

    story = []  # List to hold the flowable elements for the PDF

    # Define styles for the document
    styles = getSampleStyleSheet()

    # Update default style with built-in font
    styles['Normal'].fontName = 'Helvetica'
    styles['Normal'].fontSize = 10
    styles['Normal'].leading = 12

    # Custom styles for headings and body text
    styles.add(ParagraphStyle(
        name='CustomHeading1',
        fontName='Helvetica-Bold',
        fontSize=16,
        spaceAfter=10,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name='CustomHeading2',
        fontName='Helvetica-Bold',
        fontSize=12,
        spaceBefore=14,
        spaceAfter=4,
        textColor=colors.HexColor("#333333"),
    ))
    styles.add(ParagraphStyle(
        name="CenteredHeading1",
        fontName="Helvetica-Bold",
        fontSize=10,
        alignment=1,
        leading=15
    ))
    styles.add(ParagraphStyle(
        name='CustomBody',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
        spaceAfter=10,
    ))
    styles.add(ParagraphStyle(
        name='CustomEmphasis',
        parent=styles['CustomBody'],
        fontName='Helvetica-Bold',
        textColor=colors.HexColor("#000000"),
    ))
    # Style for table cells
    table_cell_style = ParagraphStyle(
        name='TableCell',
        fontName='Helvetica',
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#515154"),
    )
    table_cell_bold_style = ParagraphStyle(
        name='TableCellBold',
        parent=table_cell_style,
        fontName='Helvetica-Bold',
    )

    # Add Logo to the document with aspect ratio maintained
    logo = create_logo(80, 40, text_to_add["logo"])  # Adjust max dimensions as needed
    story.append(logo)
    story.append(Spacer(1, 10))  # Add space after the logo

    # Create header information with date, recipient, and sender details
    header_info = """
    Confirmation of Investor Payment Instructions<br/>
    CONFIDENTIAL
    """
    story.append(Paragraph(header_info.strip(), styles['CenteredHeading1']))

    story.append(Paragraph(f"""
    <br/><br/>
    We do not currently have wire instructions on file related to your interest in {text_to_add["fund_name"]} 
    Please enter your bank information and upload the signed form to the link provided in the e-mail associated with
this notice, by <b>{text_to_add["date"]}</b>
<br/><br/>"""))

    story.append(Paragraph(f"<b>Investor</b>: {text_to_add["legal_name"]}<br/><br/><br/>"))

    data = [

        ["", "Wire Instructions"],
        ["Bank Information", ""],
        ["Bank Name:", random.choice(banks)],
        ["ABA/SWIFT#", str(generate_SWIFT())],
        ["Account Information", ""],
        ["Beneficiary Name", text_to_add["legal_name"]],
        ["Account/IBAN #", str(generate_account_number())],
        ["",""],
        ["Address:", text_to_add["address_1"]],
        ["", ""],
        ["For Further Credit and Additional Information", ""]
    ]


    table = Table(data, colWidths=[3*inch, 3*inch])

    style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Set font size to 8 for all cells
        ('BOX', (0,0), (-1,-1), 2, colors.black),
        ('GRID', (0,1), (-1,-1), 1, colors.black)
    ])
    table.setStyle(style)

    story.append(Spacer(10, 0))  # 50 points of space horizontally
    story.append(table)


    story.append(Paragraph("""
    <br/><br/>
    If applicable, please list any additional Level Equity fund(s) the wire instructions above should also be
applied to.
<br/>"""))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*10 + "<b>Signature, Authorized Representative</b>" + "\u00A0"*30 + "<b>Contact Name for Verbal Confirmation</b>"))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*30 + "<b>Name</b>" + "\u00A0"*80 + "<b>Phone Number</b>"))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*30 + "<b>Date</b>"))


    # Build the PDF document
    doc.build(story)

def create_distribution_notice_pdf(filename, text_to_add, output_pdf_path):
    doc = fitz.open(filename)
    page = doc.load_page(0)

    logo_blocker = fitz.Rect(50, 20, 300, 60)
    top_blocker = fitz.Rect(140, 85, 500, 160)
    paragraph_blocker = fitz.Rect(50, 205, 600, 270)

    page.draw_rect(logo_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(top_blocker, color=(1, 1, 1), fill=(1, 1, 1))
    page.draw_rect(paragraph_blocker, color=(1, 1, 1), fill=(1, 1, 1))

    #logo
    #legal name
    page.insert_text((140, 95), text_to_add["legal_name"], fontsize=12, color=(0,0,0))
    #fund name
    page.insert_text((140, 117), text_to_add["fund_name"], fontsize=12, color=(0,0,0))
    #fund name re?
    page.insert_text((140, 138), text_to_add["fund_name"], fontsize=12, color=(0,0,0))
    #date
    page.insert_text((140, 158), text_to_add["date"], fontsize=12, color=(0,0,0))
    #paragraph
    paragraph_text = """
    Level Structured Capital II (GP), L.P. (“LSC II GP”) is making its first net distribution with respect to its 
    investment in Level Structured Capital II, L.P. (“LSC II”). This net distribution covers distributions of 
    investment proceeds from LSC II to LSC II (GP) from inception to date and is net of investments and 
    expenses for which capital has been called from LSC II (GP) by LSC II from inception to date
    """

    text_location = fitz.Point(40, 200)  # x=72, y=100 (1 inch margin from top-left corner)
    page.insert_text(text_location, paragraph_text, fontsize=12, fontname="helv")

    

    doc.save(output_pdf_path)
    doc.close()

# Main execution block
if __name__ == "__main__":
    # Hide the root window of Tkinter
    root = Tk()
    root.withdraw()

    # Prompt the user to select the Excel file containing the data
    print("Please select the Excel file containing the data:") 
    excel_file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )

    if not excel_file_path:
        print("No file selected. Exiting.")
        exit()

    # Read data from the selected Excel file
    df = pd.read_excel(excel_file_path, sheet_name = "Investor")

    # Prompt the user to select the output directory
    print("Please select the directory where the PDFs will be saved:")
    output_directory = filedialog.askdirectory(title="Select Output Directory")

    if not output_directory:
        print("No directory selected. Exiting.")
        exit()


    print("Select Document type: ")
    print("1. Cap Call")
    print("2. K-Document")
    print("3. Quarterly Update")
    print("4. GP Report")
    print("5. Wire Instruction Confirmation")
    print("6. Distribution Notice")

    document_type = int(input())

    while (document_type < 0 or document_type >= 7):
        print("Invalid document_type, try again")
        print("Select Document type: ")
        print("1. Cap Call")
        print("2. K-Document")
        print("3. Quarterly Update")
        print("4. GP Report")
        print("5. Wire Instruction Confirmation")
        print("6. Distribution Notice")
        document_type = int(input())


    # Get current quarter and year for filename
    now = datetime.now()
    quarter = (now.month - 1) // 3 + 1
    quarter_str = f"Q{quarter} {now.year - 1}"

    logo_path = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\aea-logo.png"

    funds = {}
    fund_names = set()

    # Iterate over each row in the DataFrame to gather all the funds
    for index, row in df.iterrows():
        fund_names.add(str(row["Fund Name"]))
    
    fund_names = list(fund_names)

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        # Updated column names based on your new input file structure
        investing_entity_name = str(row["Fund Name"])  # Fund name
        investor_code = str(row["Investor Code"])  # Investor code
        legal_name = str(row["Legal Name"])  # Investor's legal name
        address_1 = str(row["Address 1"])
        address_2 = str(row["Address 2"])
        city = str(row["City"])
        state = str(row["State"])
        zip_code = str(row["Zip"])
        country = str(row["Country"])
        tax_id = str(row["Tax ID"])
        fund_name = str(row["Fund Name"])

        #logo_path = row["Logo"]  # Path to the logo image

        # Check if the logo path is relative or absolute
        if not os.path.isabs(logo_path):
            # If relative, construct the full path relative to the Excel file's directory
            logo_path = os.path.join(os.path.dirname(excel_file_path), logo_path)

        # Verify that the logo file exists
        if not os.path.isfile(logo_path):
            print(f"Logo file '{logo_path}' not found. Skipping {legal_name}.")
            continue

        # Sanitize investor code and legal name for filenames
        investor_code_safe = sanitize_filename(investor_code)
        fund_code_safe = sanitize_filename(investor_code[4:])
        legal_name_safe = sanitize_filename(legal_name)
        output_pdf_path = ""

        text_to_add = {
            "logo" : logo_path,
            "legal_name" : legal_name,
            "footer" : f"{fund_name}, {address_1}, {city}, {state} {zip_code}",
            "date" : datetime.now().strftime("%B %d, %Y"),
            "fund_name" : fund_name,
            "address_1" : address_1
        }

        # Try-except block to catch exceptions
        try:
            #capital call
            if document_type == 1:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe} Capital Call - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_capital_call_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path
                )
                
            #k1 document
            elif document_type == 2:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe}-K1-2023.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                shutil.copy("k1-filled-flat.pdf", output_pdf_path)

            #quarterly update
            elif document_type == 3:
                #if fund has already been encountered, skip it
                if (fund_name in funds):
                    continue
                funds[fund_name] = 1

                output_pdf_name = f"{fund_code_safe} Quarterly Update Page1 - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_quarterly_update_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path
                )

                texts_with_positions = [
                    (fund_name, (30, 760))
                ]

                output_pdf_name2 = f"{fund_code_safe} Quarterly Update Page2 - {quarter_str}.pdf"
                output_pdf_path2 = os.path.join(output_directory, output_pdf_name2)
                add_multiple_texts_to_existing_pdf("resized_output.pdf", output_pdf_path2, texts_with_positions)

                output_pdf_name_final = f"{fund_name}_{fund_code_safe}_Quarterly_Update - {quarter_str}.pdf"
                output_pdf_path_final = os.path.join(output_directory, output_pdf_name_final)

                # Step 1: Open the two PDFs
                with open(output_pdf_path, "rb") as file1, open(output_pdf_path2, "rb") as file2:
                    reader1 = PyPDF2.PdfReader(file1)
                    reader2 = PyPDF2.PdfReader(file2)

                    # Step 2: Create a PdfWriter object to hold the combined PDFs
                    writer = PyPDF2.PdfWriter()

                    # Step 3: Add pages from the first PDF
                    for page_num in range(len(reader1.pages)):
                        page = reader1.pages[page_num]
                        writer.add_page(page)

                    # Step 4: Add pages from the second PDF
                    for page_num in range(len(reader2.pages)):
                        page = reader2.pages[page_num]
                        writer.add_page(page)

                    # Step 5: Write the combined PDF to a new file
                    with open(output_pdf_path_final, "wb") as output_file:
                        writer.write(output_file)

                # Step 6: Delete the original PDFs
                os.remove(output_pdf_path)
                os.remove(output_pdf_path2)
                
            #gp report
            elif document_type == 4:
                output_pdf_name = f"{fund_code_safe} GP Report - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)
                #create_gp_report_pdf(r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\templates\GP Report #4.pdf", text_to_add, output_pdf_path, all_fund_codes)

                footer = f"{investing_entity_name}, {address_1}, {city}, {state}, {zip_code}"
                create_gp_report_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path,
                    fund_names,
                    footer)

            #wire instruction confirmation
            elif document_type == 5:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe} Wire Instructions - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_wire_instruction_pdf(text_to_add, output_pdf_path)

            #distribution notice
            elif document_type == 6:
                output_pdf_name = f"{fund_name}_{investor_code_safe}_{legal_name_safe} Distribution Notice - {quarter_str}.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_distribution_notice_pdf(r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\templates\Distribution Notice #1.pdf", text_to_add, output_pdf_path)



            print(f"Generating PDF for {legal_name} at {output_pdf_path}")
            
            

        except PermissionError as e:
            print(f"Failed to write PDF for {legal_name}: {e}")
        except Exception as e:
            print(f"An error occurred while generating PDF for {legal_name}: {e}")

    print("PDF generation complete.")


