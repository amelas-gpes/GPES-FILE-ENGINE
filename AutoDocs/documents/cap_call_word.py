from docx import Document
from docx.shared import Pt, Inches

from documents.parse_excel import *

import pandas as pd 
import os

from docx2pdf import convert

def change_text(doc, position, message, font="Arial", size=10):
    # Access the specific paragraph you want to replace (e.g., the 3rd paragraph)
    paragraph = doc.paragraphs[position]

    # Clear any existing runs in the paragraph
    paragraph.clear()  # Clears the content and runs in the paragraph

    # Add a new run with the desired text
    new_run = paragraph.add_run(message)

    # Set font and size for the new run
    new_run.font.name = font  # Set font to Arial
    new_run.font.size = Pt(size)   # Set font size to 12 points

def show_doc_elements(doc):
    # Inspect paragraphs and runs
    for i, para in enumerate(doc.paragraphs):
        print(f"Paragraph {i}: {para.text}")
        for j, run in enumerate(para.runs):
            print(f"  Run {j}: {run.text}")

    # Inspect tables
    for i, table in enumerate(doc.tables):
        print(f"Table {i}:")
        for row in table.rows:
            for cell in row.cells:
                print(f"  Cell: {cell.text}")


def create_cap_call_pdf(doc, excel, fund_info, inv_info, output_directory, logo_path):

    wire_instructions = {
        "Bank Name:" : "Chase",
        "Bank Address:" : "277 Park Ave New York, NY 10172",
        "ABA Number:" : "021000021 (Domestic Wires)",
        "Account Name:" : fund_info["Fund Name"],
        "Account Number:" : "932859662",
        "SWIFT Code:" : "CHASUS33 (International Wires)"
    }

    #add logos to pages
    for section in doc.sections:
        header = section.header
        
        header_paragraph = header.paragraphs[0]
        header_paragraph.clear()

        header_paragraph.add_run().add_picture(logo_path, width=Inches(1.5))  # Adjust the size as needed

    #change header
    header = doc.tables[0]

    header.rows[2].cells[1].text = fund_info["Fund Name"]
    header.rows[4].cells[1].text = fund_info["Re"]
    header.rows[6].cells[1].text = fund_info["Notice Date"]

    header.rows[8].cells[1].text = ""
    header_fund_info = header.rows[8].cells[1].paragraphs[0].add_run(fund_info["Due Date"])
    header_fund_info.bold = True
    header_fund_info.font.name = "Arial"
    header_fund_info.font.size = Pt(10)

    message= f"""In accordance with Section 3.1 of the Amended and Restated Limited Partnership Agreement of the Fund dates {fund_info["Notice Date"]} (the "Agreement"), the Partnership is calling capital for Investments and Management Fees, and Partnership Expenses. Capitalized terms used but not defined in this notice are defined in the Agreement."""
    change_text(doc, 3, message)


    instructions = doc.tables[1]

    #Fill in wire instructions table
    for row in instructions.rows:
        if row.cells[0].text.strip():
            if str(row.cells[0].text.strip()) in wire_instructions:
                row.cells[1].text = wire_instructions[str(row.cells[0].text.strip())]


    change_text(doc, 15, fund_info["Fund Name"])


    #page 2
    re = fund_info["Re"].split()
    re = ' '.join(re[:3])
    change_text(doc, 21, "Re: " + re)


    #Fill in fund info table
    table = doc.tables[2]
    
    for row in table.rows:
        if row.cells[0].text.strip():
            if str(row.cells[0].text.strip()) in fund_info:
                row.cells[2].text = "{:,}".format(fund_info[str(row.cells[0].text.strip())])
        
    
    header.rows[0].cells[1].text = ""
    header_inv_name = header.rows[0].cells[1].paragraphs[0].add_run(inv_info["Investor Name"])
    header_inv_name.bold = True
    change_text(doc, 20, "Investor: " + inv_info["Investor Name"])
    #Fill in the info
    instructions.rows[7].cells[1].text = inv_info["Investor Name"]
    message = f"""Your portion of the call is ${inv_info["Total Amount Due"]:,} and is due on {fund_info["Due Date"]}.  Please send your payment by wire transfer in accordance with the instructions provided below."""
    change_text(doc, 7, message)
    for row in table.rows:
        if row.cells[0].text.strip():
            if str(row.cells[0].text.strip()) in inv_info:
                row.cells[5].text = "{:,}".format(inv_info[str(row.cells[0].text.strip())])
    print(output_directory)
    investor_docx = os.path.join(output_directory, f"{inv_info['Investor Name']}.docx")
    investor_pdf = os.path.join(output_directory, f"{inv_info['Investor Name']}.pdf")
    doc.save(investor_docx)
    # Convert the Word document to PDF
    convert(investor_docx, investor_pdf)
    os.remove(investor_docx)

