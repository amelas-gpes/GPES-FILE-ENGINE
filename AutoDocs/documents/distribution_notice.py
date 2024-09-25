from documents.utils import *

def create_distribution_notice_pdf(text_to_add, output_pdf_path):
    doc = SimpleDocTemplate(
        output_pdf_path, pagesize=letter,
        rightMargin=0.75 * inch, leftMargin=0.75 * inch,
        topMargin=0.25 * inch, bottomMargin=0.25 * inch
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
    styles.add(ParagraphStyle(
        name="Footer_small",
        fontName="Helvetica",
        fontSize=8,
        leading=8
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
    #story.append(logo)
    #story.append(Spacer(1, 10))  # Add space after the logo

    # Sample data for the table with an image in one cell and text in another
    top_data = [
        [logo, Paragraph("<b>CONFIDENTIAL</b>", styles["Normal"])]
    ]

    # Create a table
    top_table = Table(top_data, colWidths=[5*inch, 1.5*inch], hAlign="LEFT")

    # Apply a basic style to the table
    top_style = TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Align cells vertically in the middle
    ])
    top_table.setStyle(top_style)

    story.append(Spacer(1, 10))  # Add space before the table
    story.append(top_table)

    header_data = [
        ["To", Paragraph(f"<b>{text_to_add['legal_name']}</b>", styles["Normal"])],
        ["From:", f"{text_to_add['fund_name']}"],
        ["RE:", f"{text_to_add['fund_name']}"],
        ["Date:", f"{text_to_add['date']}"]
    ]

    header_table = Table(header_data, colWidths=[inch, 3*inch], hAlign="LEFT")

    header_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 10),  # Set font size to 8 for all cells
    ])
    header_table.setStyle(header_style)

    story.append(Spacer(1, 10))  # Add space before the table
    story.append(header_table)

    story.append(Paragraph("""<br/><b>Net Distribution #1: Distributable Proceeds; net of contributions for Investments, Organizational
Expenses, and Partnership Expenses</b>
"""))

    story.append(DividerLine(doc.width))

    story.append(Paragraph(f"""<br/>{text_to_add["fund_name"]} is making its first net distribution with respect to its
investment in {text_to_add["fund_name"]}. This net distribution covers distributions of
investment proceeds from {text_to_add["fund_name"]} to {text_to_add["fund_name"]} from inception to date and is net of investments and
expenses for which capital has been called from {text_to_add["fund_name"]} by {text_to_add["fund_name"]} from inception to date.<br/><br/>"""))

    num = random.randrange(10,99)

    body_data = [
        ["", "", Paragraph("<b>Current</b>", styles["Normal"]), "", Paragraph("<b>Cumulative</b>", styles["Normal"])],
        [Paragraph("<b>Allocation of your Distributable Proceeds:</b>", styles["Normal"]), "$", "-", "$", "-"],
        ["  Committed Capital", "$", "-", "$", "-"],
        ["  Carried Interest", "", f"{num}", "", f"{num}"],
        ["Total Distribution", "$", f"{num}", "$", f"{num}"],
        ["","","","",""],
        [Paragraph("<b>Capital Call</b>", styles["Normal"]), "$", "-", "$", "-"],
        ["  Your Share of GP Gross Contributions", "$", "-", "$", "-"],
        ["  Your Share of GP Deemed Contributions", "", "-", "", "-"],
        ["  Your Share of GP Standalone Expenses", "", "-", "", "-"],
        ["Total Capital Call", "$", "-", "$", "-"],
        ["","","","",""],
        ["Net Distribution", "$", f"{num}", "", ""],
        ["Less: Equalization Interest", "$", "-", "", ""],
        [Paragraph("<b>Net Amount Payable</b>", styles["Normal"]), "$", f"{num}", "", ""],
        [Paragraph("Total Estimated Taxable Income<super>1</super>", styles['Normal']), "", "", "$", f"{num}"],
        ["", "", "", "", ""],
        ["","","","",""],
        [Paragraph("<b>Your Remaining Commitment after this Net Distribution</b>", styles["Normal"]), "", "", "", Paragraph("<b>Cumulative</b>", styles["Normal"])],
        ["  Capital Commitment", "", "", "$", "-"],
        ["  Cumulative Contributions", "", "", "$", "-"],
        ["  Remaining Commitment", "", "", "$", "-"],
    ]

    body_table = Table(body_data, colWidths=[4*inch, 0.5*inch, inch, 0.5*inch, inch], hAlign="LEFT")

    body_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 10),  # Set font size to 8 for all cells
        ('ROWHEIGHTS', (0, 0), (-1, -1), 1),  # Set row heights
        ('TOPPADDING', (0, 0), (-1, -1), 2),  # Reduce top padding
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Reduce bottom padding
    ])
    body_table.setStyle(body_style)

    story.append(Spacer(1, 10))  # Add space before the table
    story.append(body_table)

    story.append(Spacer(1, 30))
    story.append(DividerLine(doc.width))

    story.append(Spacer(1,10))
    story.append(Paragraph("""<super>1</super>Tax information is an estimate only. Final amounts will be reported on your 2022 and 2023 Schedule K-1s. You should consult
your tax adviser with respect to any federal or state tax consequences. Any information provided is not intended to be used,
and cannot be used, to avoid penalties imposed under the Internal Revenue Service. Please refer to your Semi-Annual partner
capital account statement for your share of current year GAAP capital gain/(loss).""", styles["Footer_small"]))
    story.append(Spacer(1,10))

    story.append(DividerLine(doc.width))

    footer_data = [
        [f"{text_to_add['address_1']}", "CONFIDENTIAL"],
        [f"{text_to_add['state']}, {text_to_add['city']} {text_to_add['zip_code']}", ""],
    ]
    footer_table = Table(footer_data, colWidths=[6*inch, inch], hAlign="LEFT")

    footer_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.white),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),  # Using Helvetica for all cells
        ('FONTSIZE', (0,0), (-1,-1), 8),  # Set font size to 8 for all cells
        ('TOPPADDING', (0, 0), (-1, -1), 1),  # Reduce top padding
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),  # Reduce bottom padding
    ])
    footer_table.setStyle(footer_style)

    story.append(Spacer(1, 10))  # Add space before the table
    story.append(footer_table)



    # Build the PDF document
    doc.build(story)


