import PyPDF2
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter

# Step 1: Open the original PDF
with open("quarterly_report.pdf", "rb") as file:
    reader = PdfReader(file)
    writer = PdfWriter()

    # Step 2: Loop through each page in the original PDF
    for page_num in range(len(reader.pages)):
        original_page = reader.pages[page_num]

        # Get the original page size
        original_width = original_page.mediabox.width
        original_height = original_page.mediabox.height

        # New page dimensions (Letter size in points)
        new_width, new_height = letter

        # Step 3: Calculate scale factors to fit the content into the new page size
        scale_x = new_width / original_width
        scale_y = new_height / original_height

        # Use the smaller scaling factor to maintain the aspect ratio
        scale_factor = min(scale_x, scale_y)

        # Step 4: Scale the original page content
        original_page.scale_by(scale_factor)

        # Step 5: Center the scaled content on the new page size
        original_page.mediabox.upper_right = (new_width, new_height)

        # Step 6: Add the scaled page to the PdfWriter
        writer.add_page(original_page)

    # Step 7: Write the resized PDF to a new file
    with open("resized_output.pdf", "wb") as output_pdf:
        writer.write(output_pdf)
