from documents.utils import *

def create_k1_document_pdf(fund_name, legal_name, output_pdf_path):
    __location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    input_pdf = __location__ + "\k1_template.pdf"
    texts_with_positions = [
        (fund_name + " 1456 Sierra Ridge Drive Fresno CA 93711", (40, 580)),
        (legal_name + " 8973 Elliott Stream South Shawnchester CA 97695", (40, 480)),
    ]


    add_multiple_texts_to_existing_pdf(input_pdf, output_pdf_path, texts_with_positions)

