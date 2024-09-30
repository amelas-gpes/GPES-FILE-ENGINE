from documents.utils import *

"""
Generates a capital call PDF document with the given filename and content.

Parameters:
- filename: Name of the output PDF file.
- investing_entity_name: Name of the investing entity (fund).
- legal_name: Legal name of the investor.
- image_path: Path to the logo image file.
"""
def create_capital_call_pdf(filename, investing_entity_name, legal_name, image_path, color="#515154"):
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
    placement_style = ParagraphStyle(
        name="sample_style",
        parent=table_cell_style,
        textColor = colors.HexColor(color)
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
    <b>Date:</b> <font color="{color}">{current_date}</font><br/>
    <b>To:</b> <font color="{color}">{legal_name}</font><br/>
    <b>From:</b> <font color="{color}">{investing_entity_name}</font>
    """
    story.append(Paragraph(header_info.strip(), styles['CustomBody']))

    # Divider line
    story.append(Spacer(1, 6))
    story.append(DividerLine(doc.width))
    story.append(Spacer(1, 6))

    # Company Description
    company_description = f"""
    <font color="{color}">{investing_entity_name}</font> is a private equity investment fund focused on identifying and nurturing high-potential growth companies across various sectors. With a proven track record of strategic investments and value creation, our fund aims to generate superior returns for our limited partners while fostering innovation and sustainable growth in our portfolio companies.
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
            Paragraph("${:,}".format(total_commitment), placement_style)
        ],
        [
            Paragraph("Prior Capital Contributions", table_cell_style),
            Paragraph("${:,}".format(prior_capital_contributions), placement_style)
        ],
        [
            Paragraph("Unfunded Commitment", table_cell_style),
            Paragraph("${:,}".format(unfunded_commitment), placement_style)
        ],
        [
            Paragraph("Current Capital Call", table_cell_style),
            Paragraph("${:,}".format(current_capital_call), placement_style)
        ],
        [
            Paragraph("Remaining Unfunded Commitment", table_cell_style),
            Paragraph("${:,}".format(remaining_unfunded_commitment), placement_style)
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
    • New Portfolio Company Investment: <b><font color="{0}">${1:,}</font></b><br/>
    • Follow-on Investment: <b><font color="{0}">${2:,}</font></b><br/>
    • Fund Expenses: <b><font color="{0}">${3:,}</font></b>
    """.format(
        color,
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
    <b>Bank:</b> <font color="{color}">{generate_bank_name()}</font><br/>
    <b>Account:</b> <font color="{color}">{investing_entity_name}</font><br/>
    <b>Account Number:</b> <b><font color="{color}">{account_number}</font></b><br/>
    <b>ABA:</b> <b><font color="{color}">{ABA}</font></b><br/>
    <b>SWIFT:</b> <font color="{color}">{SWIFT}</font>
    """
    story.append(Paragraph(payment.strip(), styles['CustomBody']))

    # Due Date Section with emphasis
    due_date = (datetime.now() + pd.DateOffset(days=7)).strftime("%B %d, %Y")
    story.append(Spacer(1, 10))
    due_date_text = f"<b>Due Date:</b> <b><font color='{color}'>{due_date}</font></b>"
    story.append(Paragraph(
        due_date_text,
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
    <b>Contact:</b> <font color="{color}">{contact_name}</font>, <font color="{color}">{contact_role}</font><br/>
    <font color="{color}">{contact_name.replace(' ', '.').lower()}@{investing_entity_name.lower().replace(' ', '')}.com</font> | <font color="{color}">{format_phone_number(random.randrange(100, 1000), random.randrange(100, 1000), random.randrange(1000, 10000))}</font>
    """
    story.append(Paragraph(contact.strip(), styles['CustomBody']))

    # Build the PDF document
    doc.build(story)