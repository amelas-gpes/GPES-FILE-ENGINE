from documents.capital_call import *
from documents.quarterly_update import *
from documents.gp_report import *
from documents.wire_instruction import *
from documents.distribution_notice import *
from documents.k1_document import *

from documents.utils import *

import tkinter as tk
from tkinter import filedialog

def select_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        label_file.config(text=f"Selected: {file_path}")
        selected_file.set(file_path)

def select_logo():
    logo_path = filedialog.askopenfilename(
        title="Select an Image",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp")],  # Filter image file types
    )
    if logo_path:
        label_logo.config(text=f"Selected: {logo_path}")
        selected_logo.set(logo_path)

def select_directory():
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if output_directory:
        label_dir.config(text=f"Selected: {output_directory}")
        selected_directory.set(output_directory)

def option_selected(choice):
    label_option.config(text=f"Selected option: {choice}")


def submit_action():
    excel_file_path = selected_file.get()
    logo_path = selected_logo.get()
    output_directory = selected_directory.get()
    option = selected_option.get()

    if excel_file_path == "No file selected":
        return
    if logo_path == "No logo selected":
        return
    if output_directory == "No directory selected":
        return
    if option == "No option selected":
        return

    submit_label.config(text="Generating...")

    # Read data from the selected Excel file
    df = pd.read_excel(excel_file_path)


    # Get current quarter and year for filename
    now = datetime.now()
    quarter = (now.month - 1) // 3 + 1
    quarter_str = f"Q{quarter} {now.year - 1}"

    funds = {}
    fund_names = set()

    # Iterate over each row in the DataFrame to gather all the funds
    for index, row in df.iterrows():
        fund_names.add(str(row["Fund Name"]))
    
    fund_names = list(fund_names)


    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        investing_entity_name = str(row["Fund Name"]) 
        investor_code = str(row["Investor Code"])  
        legal_name = str(row["Legal Name"]) 
        address_1 = str(row["Address 1"])
        address_2 = str(row["Address 2"])
        city = str(row["City"])
        state = str(row["State"])
        zip_code = str(row["Zip"])
        country = str(row["Country"])
        tax_id = str(row["Tax ID"])
        fund_name = str(row["Fund Name"])

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
            if option == "Capital Call":
                output_pdf_name = f"{investor_code_safe}_{legal_name_safe} - {fund_name} - Capital Call.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_capital_call_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path
                )

            
            #k1 document
            elif option == "K1 Document":
                output_pdf_name = f"{investor_code_safe}_{legal_name_safe} - {fund_name} - K1.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_k1_document_pdf(fund_name, legal_name, output_pdf_path)

            #quarterly update
            elif option == "Quarterly Update":
                #if fund has already been encountered, skip it
                if (fund_name in funds):
                    continue
                funds[fund_name] = 1

                output_pdf_name = f"{fund_code_safe} Quarterly Update Page - {quarter_str}.pdf"
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

                output_pdf_name_final = f"{fund_code_safe}_{fund_name} - Quarterly Update.pdf"
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
            elif option == "GP Report":
                output_pdf_name = f"{fund_code_safe}_{fund_name} - GP Report.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                footer = f"{investing_entity_name}, {address_1}, {city}, {state}, {zip_code}"
                create_gp_report_pdf(
                    output_pdf_path,
                    investing_entity_name,
                    legal_name,
                    logo_path,
                    fund_names,
                    footer)

            #wire instruction confirmation
            elif option == "Wire Instruction":
                output_pdf_name = f"{investor_code_safe}_{legal_name_safe} - {fund_name} - Wire Instructions.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                create_wire_instruction_pdf(text_to_add, output_pdf_path)

            #distribution notice
            elif option == "Distribution Notice":
                output_pdf_name = f"{investor_code_safe}_{legal_name_safe} - {fund_name} - Distribution Notice.pdf"
                output_pdf_path = os.path.join(output_directory, output_pdf_name)

                text_to_add = {
                        "logo" : logo_path,
                        "legal_name" : legal_name_safe,
                        "date" : datetime.now().strftime("%B %d, %Y"),
                        "fund_name" : fund_name,
                        "address_1" : address_1,
                        "state" : state,
                        "city" : city,
                        "zip_code" : zip_code,
                    }

                create_distribution_notice_pdf(text_to_add, output_pdf_path)

            
            print(f"Generating PDF for {legal_name} at {output_pdf_path}")
                
        except PermissionError as e:
            print(f"Failed to write PDF for {legal_name}: {e}")
        except Exception as e:
            print(f"An error occurred while generating PDF for {legal_name}: {e}")
        
        
    print("PDF generation complete.")
    submit_label.config(text="Done!")


# Main execution block
if __name__ == "__main__":
    
    # Create the root window
    root = tk.Tk()
    root.title("AutoDocs")

    # Set the window size
    root.geometry("600x650")
    root.config(bg="#f0f0f0")  # Background color

    # Variables to store user selections
    selected_file = tk.StringVar(root, value="No file selected")
    selected_logo = tk.StringVar(root, value="No logo selected")
    selected_directory = tk.StringVar(root, value="No directory selected")

    # Define font and styling options
    font_label = ("Helvetica", 12)
    font_button = ("Helvetica", 12, "bold")

    # Create a frame for file selection
    frame_file = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
    frame_file.pack(padx=20, pady=20, fill="x")

    label_file_title = tk.Label(frame_file, text="File Selection", font=font_label, bg="#e6e6e6")
    label_file_title.pack(anchor="w")

    button_file = tk.Button(frame_file, text="Select File", command=select_file, font=font_button, bg="#4CAF50", fg="white")
    button_file.pack(side="left", padx=10, pady=5)

    label_file = tk.Label(frame_file, textvariable=selected_file, bg="#e6e6e6", font=font_label)
    label_file.pack(side="left", padx=10)

    # Create a frame for file selection
    frame_logo = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
    frame_logo.pack(padx=20, pady=20, fill="x")

    label_logo_title = tk.Label(frame_logo, text="Logo Selection", font=font_label, bg="#e6e6e6")
    label_logo_title.pack(anchor="w")

    button_logo = tk.Button(frame_logo, text="Select Logo", command=select_logo, font=font_button, bg="#4CAF50", fg="white")
    button_logo.pack(side="left", padx=10, pady=5)

    label_logo = tk.Label(frame_logo, textvariable=selected_logo, bg="#e6e6e6", font=font_label)
    label_logo.pack(side="left", padx=10)

    # Create a frame for directory selection
    frame_dir = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
    frame_dir.pack(padx=20, pady=20, fill="x")

    label_dir_title = tk.Label(frame_dir, text="Directory Selection", font=font_label, bg="#e6e6e6")
    label_dir_title.pack(anchor="w")

    button_dir = tk.Button(frame_dir, text="Select Output Directory", command=select_directory, font=font_button, bg="#4CAF50", fg="white")
    button_dir.pack(side="left", padx=10, pady=5)

    label_dir = tk.Label(frame_dir, textvariable=selected_directory, bg="#e6e6e6", font=font_label)
    label_dir.pack(side="left", padx=10)



    # Create a frame for the option menu
    frame_option = tk.Frame(root, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
    frame_option.pack(padx=20, pady=20, fill="x")

    label_option_title = tk.Label(frame_option, text="Select an Option", font=font_label, bg="#e6e6e6")
    label_option_title.pack(anchor="w")

    options = ["Capital Call", "Distribution Notice", "GP Report", "Wire Instruction", "Quarterly Update", "K1 Document"]
    selected_option = tk.StringVar(root)
    selected_option.set(options[0])

    option_menu = tk.OptionMenu(frame_option, selected_option, *options, command=option_selected)
    option_menu.config(font=font_button, bg="#4CAF50", fg="white")
    option_menu.pack(side="left", padx=10, pady=5)

    label_option = tk.Label(frame_option, text="No option selected", bg="#e6e6e6", font=font_label)
    label_option.pack(side="left", padx=10)

    # Create a Submit button
    submit_button = tk.Button(root, text="Submit", command=submit_action, font=font_button, bg="#008CBA", fg="white")
    submit_button.pack(pady=20)

    # Label to show the result after submission
    submit_label = tk.Label(root, text="", bg="#f0f0f0", font=font_label)
    submit_label.pack(pady=10)

    # Run the application
    root.mainloop()


    
