from documents.utils import *

"""
Generates a GP report PDF document with the given filename and content.

Parameters:
- filename: Name of the output PDF file.
- investing_entity_name: Name of the investing entity (fund).
- legal_name: Legal name of the investor.
- image_path: Path to the logo image file.
- fund_names: all the names of the funds
- footer: fund name, address, state, ny, zipcode
"""
def create_gp_report_pdf(filename, investing_entity, legal_name, image_path, fund_names, footer):
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

    data = [
    ]

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

    #generate distinct fund names 
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
    "-", "-", "", "",
    "", "", "",
    "", "-", "", "", "-", "", 
    "", ""]

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

        row.append(Paragraph(row_titles[index], table_cell_style))
        row.append(Paragraph(currency_col[index], table_cell_style))

        for j in range(len(fund_info)):
            if index >= len(fund_info[j]):
                row.append(Paragraph("", table_cell_style))
                continue
            row.append(Paragraph(fund_info[j][index], table_cell_style))
        
        row.append(Paragraph(total[index], table_cell_style))

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