from documents.utils import *

"""
Generates a wire instruction confirmation PDF document.

Parameters:
- text_to_add: Dictionary containing text elements for the PDF.
- output_pdf_path: Path to save the generated PDF.
"""
def create_wire_instruction_pdf(text_to_add, output_pdf_path):
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
        fontSize=8,
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
    We do not currently have wire instructions on file related to your interest in {text_to_add["fund_name"]}. 
    Please enter your bank information and upload the signed form to the link provided in the e-mail associated with
    this notice, by <b>{text_to_add["date"]}</b>
    <br/><br/>""", styles['CustomBody']))

    story.append(Paragraph(f"<b>Investor</b>: {text_to_add['legal_name']}<br/><br/><br/>", styles['CustomBody']))

    data = [

        ["", "Wire Instructions"],
        ["Bank Information", ""],
        ["Bank Name:", generate_bank_name()],
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
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Set font size to 8 for all cells
        ('BOX', (0,0), (-1,-1), 2, colors.black),
        ('GRID', (0,1), (-1,-1), 1, colors.black)
    ])
    table.setStyle(style)

    story.append(Spacer(1, 10))  # Add space before the table
    story.append(table)

    story.append(Paragraph("""
    <br/><br/>
    If applicable, please list any additional funds the wire instructions above should also be
    applied to.
    <br/>""", styles['CustomBody']))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*10 + "<b>Signature, Authorized Representative</b>" + "\u00A0"*30 + "<b>Contact Name for Verbal Confirmation</b>", styles['CustomBody']))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*30 + "<b>Name</b>" + "\u00A0"*80 + "<b>Phone Number</b>", styles['CustomBody']))

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Paragraph("\u00A0"*30 + "<b>Date</b>", styles['CustomBody']))

    # Build the PDF document
    doc.build(story)
