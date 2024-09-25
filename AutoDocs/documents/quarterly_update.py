from documents.utils import *

"""
Generates a capital call PDF document with the given filename and content.

Parameters:
- filename: Name of the output PDF file.
- investing_entity_name: Name of the investing entity (fund).
- legal_name: Legal name of the investor.
- image_path: Path to the logo image file.
"""
def create_quarterly_update_pdf(filename, investing_entity_name, legal_name, image_path):
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