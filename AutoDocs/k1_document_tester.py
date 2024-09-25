from documents.k1_document import *

if __name__ == "__main__":
    output_pdf_path = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\K1_DOCUMENT_TEST.pdf"

    text_to_add = {
        "logo" : r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\doc_generator\aea-logo.png",
        "legal_name" : "legal_name_safe",
        "date" : datetime.now().strftime("%B %d, %Y"),
        "fund_name" : "fund_name",
        "address_1" : "address_1",
        "state" : "state",
        "city" : "city",
        "zip_code" : "zip_code",
    }

    create_distribution_notice_pdf(text_to_add, output_pdf_path)
