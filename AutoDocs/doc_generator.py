from documents.capital_call import *
from documents.quarterly_update import *
from documents.gp_report import *
from documents.wire_instruction import *
from documents.distribution_notice import *
from documents.k1_document import *
from documents.cap_call_word import *

from documents.utils import *

import tkinter as tk
from tkinter import filedialog

import os
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white
from reportlab.lib.units import inch
from io import BytesIO

from pdf_viewer import *

class MainApp(tk.Tk):
    def select_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.label_file.config(text=f"Selected: {file_path}")
            self.selected_file.set(file_path)

            allocation = pd.read_excel(file_path, sheet_name = "Allocation", skiprows=1)
            index = 5
            while (str(allocation.at[index, "Partner Name"]) != "nan"):
                self.investors.add(str(allocation.at[index, "Partner Name"]))
                index += 1
            self.checked_investors = list(self.investors)
            
            """
            df = pd.read_excel(file_path, sheet_name = "Investor")
            investor_col = df["Legal Name"]
            for i in investor_col:
                self.investors.add(i)
            self.checked_investors = list(self.investors)
            """



    def select_logo(self):
        logo_path = filedialog.askopenfilename(
            title="Select an Image",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp")],  # Filter image file types
        )
        if logo_path:
            self.label_logo.config(text=f"Selected: {logo_path}")
            self.selected_logo.set(logo_path)

    def select_directory(self):
        output_directory = filedialog.askdirectory(title="Select Output Directory")
        if output_directory:
            self.label_dir.config(text=f"Selected: {output_directory}")
            self.selected_directory.set(output_directory)

    def option_selected(self, choice):
        self.label_option.config(text=f"Selected option: {choice}")

    def submit_action(self):
        excel_file_path = self.selected_file.get()
        logo_path = self.selected_logo.get()
        output_directory = self.selected_directory.get()
        option = self.selected_option.get()

        first_name = self.first_name.get()
        last_name = self.last_name.get()
        email = self.email.get()

        start_delim = self.start_delimiter.get()
        end_delim = self.end_delimiter.get()
        
        if self.is_sample:
            output_directory = "./sample/"


        if excel_file_path == "No file selected":
            print("No file selected")
            return
        if logo_path == "No logo selected":
            print("No logo selected")
            return
        if output_directory == "No directory selected":
            print("No directory selected")
            return
        if option == "No option selected":
            print("No option selectes")
            return
        if first_name == "":
            print("No first name given")
            return
        if last_name == "":
            print("No last name given")
            return
        if email == "":
            print("No email given")
            return
        

        color = "#515154"
        if self.is_sample:
            color = "#ff4e28"

        fund_info = parse_excel(excel_file_path)
        fund_info["Start Delim"] = start_delim
        fund_info["End Delim"] = end_delim


        #go through each investors and create pdfs for them
        allocation = pd.read_excel(excel_file_path, sheet_name = "Allocation", skiprows=1)
        index = 5
        while (str(allocation.at[index, "Partner Name"]) != "nan"):
            if (str(allocation.at[index, "Partner Name"]) not in self.checked_investors):
                index += 1
                continue
            
            #Gather info
            inv_info = dict()

            inv_info["Investor Name"] = str(allocation.at[index, "Partner Name"])

            #investments = int(allocation.at[index, "Investment #1"])
            inv_info["Investment"] = int(allocation.at[index, "Investment #1"])

            #mgmt_fees = int(allocation.at[index, "Gross Mgmt Fee"])
            inv_info["Management Fees"] = int(allocation.at[index, "Gross Mgmt Fee"])

            #pshp_exp = int(allocation.at[index, "Pshp Exp"])
            inv_info["Partnership Expenses"] = int(allocation.at[index, "Pshp Exp"])

            total_amnt = int(allocation.at[index, "Net Amount Due / (Payable)"])
            inv_info["Total Amount Due"] = int(allocation.at[index, "Net Amount Due / (Payable)"])

            cap_commit = int(allocation.at[index, "Commitment"])
            inv_info["Capital Commitment"] = int(allocation.at[index, "Commitment"])

            cum_cap_contributions = int(allocation.at[index, "LTD Ending Contributions"])
            inv_info["Cumulative Capital Contributions"] = int(allocation.at[index, "LTD Ending Contributions"])

            rem_cap_commit = int(allocation.at[index, "Ending Remaining  Commitment"])
            inv_info["Remaining Capital Commitment"] = int(allocation.at[index, "Ending Remaining  Commitment"])

            commit_subj_mgmt_fee = int(allocation.at[index, "LP Commitment"])
            inv_info["Commitment subject to Management Fee"] = int(allocation.at[index, "LP Commitment"])

            mgmt_fee = int(allocation.at[index, "Gross Mgmt Fee"])
            inv_info["Management Fee (1/1/2022 - 3/31/2022)"] = int(allocation.at[index, "Gross Mgmt Fee"])

            mgmt_fee_reduct = int(allocation.at[index, "Mgmt Fee - Offsets"])
            inv_info["Management Fee Reduction"] = int(allocation.at[index, "Mgmt Fee - Offsets"])

            total_mgmt_fee = int(allocation.at[index, "Total Mgmt Fee"])
            inv_info["Total Management Fee, net"] = int(allocation.at[index, "Total Mgmt Fee"])

            inv_info["First Name"] = first_name
            inv_info["Last Name"] = last_name
            inv_info["Email"] = email

            index += 1
            
            if option == "Capital Call":
                doc = Document(r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\git\GPES-FILE-ENGINE\AutoDocs\documents\cap_call_template.docx")
                
                output_pdf_name = inv_info["Investor Name"]
                if self.is_sample:
                    output_pdf_name = "sample"

                output_path = os.path.join(output_directory, f"{output_pdf_name}")
                if not self.is_sample:
                    self.files_list.append(output_path+".pdf")

                create_cap_call_pdf(doc, excel_file_path, fund_info, inv_info, output_path, logo_path)

            if self.is_sample:
                self.sample_ready = True
                break

        if (self.bulk_choice.get()):
            self.merge_pdfs_in_folder(output_directory, "bulk.pdf")

        self.files_list = []
        print("PDF generation complete.")

    def merge_pdfs_in_folder(self, folder_path, output_filename):
        # Create a PdfMerger object
        merger = PdfMerger()

        # Loop through all the files in the folder
        for filename in sorted(os.listdir(folder_path)):
            if filename.endswith(".pdf"):
                file_path = os.path.join(folder_path, filename)
                # Append each PDF to the merger
                with open(file_path, 'rb') as pdf_file:
                    merger.append(pdf_file)
                if (not self.split_choice.get()):
                    os.remove(file_path)

        # Write the merged PDF to a new file
        output_path = os.path.join(folder_path, output_filename)
        with open(output_path, 'wb') as output_file:
            merger.write(output_file)

        # Close the merger
        merger.close()

        print(f"All PDFs merged into {output_filename}")


    def __init__(self):
        super().__init__()
        self.title("AutoDocs")
        # Set the window size
        self.geometry("1280x1000")
        self.config(bg="#f0f0f0")  # Background color

        # Variables to store user selections
        self.selected_file = tk.StringVar(self, value="No file selected")
        self.selected_logo = tk.StringVar(self, value="No logo selected")
        self.selected_directory = tk.StringVar(self, value="No directory selected")
        self.first_name = tk.StringVar(self, value="")
        self.last_name = tk.StringVar(self, value="")
        self.email = tk.StringVar(self, value="")

        self.label_file = ""
        self.label_logo = ""
        self.label_dir = ""

        self.investors = set()
        self.checked_investors = list(self.investors)

        self.is_sample = False

        # Create a StringVar to hold the choice (split or bulk)
        self.split_choice = tk.IntVar(value=1)
        self.bulk_choice = tk.IntVar(value=0)
        self.start_delimiter = tk.StringVar(self, value = "<start>")
        self.end_delimiter = tk.StringVar(self, value = "<end>")
        self.output_file = tk.StringVar(self, value = "bulk.pdf")

        self.files_list = []

        # Define font and styling options
        self.font_label = ("Helvetica", 12)
        self.font_button = ("Helvetica", 12, "bold")

        # Create a container to hold the frames (pages)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Dictionary to hold pages
        self.frames = {}

        # Initialize all the pages
        for F in (InputPage, OutputPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Show the home page initially
        self.show_frame("InputPage")

    def show_frame(self, page_name):
        # Bring the frame to the front
        frame = self.frames[page_name]
        frame.tkraise()
        

class InputPage(tk.Frame):
    def __init__(self, parent, controller):
        def on_next():
            controller.show_frame("OutputPage")

        super().__init__(parent)
        self.controller = controller
        
        label = tk.Label(self, text="GPES FileGen", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)

        # Frame for file selection
        frame_file = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_file.pack(padx=20, pady=20, fill="x")

        label_file_title = tk.Label(frame_file, text="Select Excel File", font=controller.font_label, bg="#e6e6e6")
        label_file_title.pack(anchor="w")

        button_file = tk.Button(frame_file, text="Select File", command=controller.select_file, font=controller.font_button, bg="#4CAF50", fg="white")
        button_file.pack(side="left", padx=10, pady=5)

        controller.label_file = tk.Label(frame_file, textvariable=controller.selected_file, bg="#e6e6e6", font=controller.font_label)
        controller.label_file.pack(side="left", padx=10)

        # Frame for logo selection
        frame_logo = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_logo.pack(padx=20, pady=20, fill="x")

        label_logo_title = tk.Label(frame_logo, text="Select Logo To Be Placed On Documents", font=controller.font_label, bg="#e6e6e6")
        label_logo_title.pack(anchor="w")

        button_logo = tk.Button(frame_logo, text="Select Logo", command=controller.select_logo, font=controller.font_button, bg="#4CAF50", fg="white")
        button_logo.pack(side="left", padx=10, pady=5)

        controller.label_logo = tk.Label(frame_logo, textvariable=controller.selected_logo, bg="#e6e6e6", font=controller.font_label)
        controller.label_logo.pack(side="left", padx=10)

        # Frame for the option menu
        frame_option = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_option.pack(padx=20, pady=20, fill="x")

        label_option_title = tk.Label(frame_option, text="Select An Option", font=controller.font_label, bg="#e6e6e6")
        label_option_title.pack(anchor="w")

        options = ["Capital Call", "Distribution Notice"]
        controller.selected_option = tk.StringVar(self)
        controller.selected_option.set(options[0])

        option_menu = tk.OptionMenu(frame_option, controller.selected_option, *options, command=controller.option_selected)
        option_menu.config(font=controller.font_button, bg="#4CAF50", fg="white")
        option_menu.pack(side="left", padx=10, pady=5)

        controller.label_option = tk.Label(frame_option, text="No option selected", bg="#e6e6e6", font=controller.font_label)
        controller.label_option.pack(side="left", padx=10)


        # Frame for contact info
        frame_contacts = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_contacts.pack(padx=20, pady=20, fill="x")

        label_contacts = tk.Label(frame_contacts, text="Contact Info", font=controller.font_label, bg="#e6e6e6")
        label_contacts.pack(anchor="w")

        # First Name
        label_first_name = tk.Label(frame_contacts, text="First Name: ", font=controller.font_label, bg="#e6e6e6")
        label_first_name.pack(anchor="w", pady=(10, 0))  # Place the label at the top-left

        first_name_entry = tk.Entry(frame_contacts, textvariable=controller.first_name)
        first_name_entry.pack(anchor="w", fill="x")  # Fill horizontally

        # Last Name
        label_last_name = tk.Label(frame_contacts, text="Last Name: ", font=controller.font_label, bg="#e6e6e6")
        label_last_name.pack(anchor="w", pady=(10, 0))  # Space between first and last name

        last_name_entry = tk.Entry(frame_contacts, textvariable=controller.last_name)
        last_name_entry.pack(anchor="w", fill="x")  # Fill horizontally

        # Email
        label_email = tk.Label(frame_contacts, text="Email: ", font=controller.font_label, bg="#e6e6e6")
        label_email.pack(anchor="w", pady=(10,0))

        email_entry = tk.Entry(frame_contacts, textvariable=controller.email)
        email_entry.pack(anchor="w", fill="x")


        # Frame for the investors input
        frame_investors = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_investors.pack(padx=20, pady=20, fill="x")

        self.select_all = True
        investors_button = tk.Button(frame_investors, text="Select Investors", command=self.show_investors)
        investors_button.pack()
        
        # Button to go to next page
        button1 = tk.Button(self, text="Next", command= on_next)
        button1.pack()

    def show_investors(self):
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_mouse_wheel(event):
            # Cross-platform scrolling using event.delta
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def update_list():
            # Keep only the items that are checked (i.e., where investor_value.get() == 1)
            self.controller.checked_investors = [item for i, item in enumerate(item_list, 1) if vars[i].get() == 1]

        def select_all():
            if (vars[0].get() == 1):
                self.controller.checked_investors = self.controller.investors
                for i in checkboxes:
                    i.select()
                self.select_all = True
            else:
                self.controller.checked_investors = []
                for i in checkboxes:
                    i.deselect()
                self.select_all = False

        def create_checkboxes():
            # Create checkbox for selecting all
            select_value = tk.IntVar(value=self.select_all)
            vars.append(select_value)
            checkbox = tk.Checkbutton(scrollable_frame, text="Select All", variable=select_value, command=select_all)
            checkbox.pack(anchor="w", pady=5)
            checkboxes.append(checkbox)

            # Create checkboxes for each item in the list
            for i, item in enumerate(item_list):
                investor_value = tk.IntVar(value=(item in self.controller.checked_investors))  # Variable to track the state of the checkbox
                vars.append(investor_value)
                checkbox = tk.Checkbutton(scrollable_frame, text=item, variable=investor_value, command=update_list)
                checkbox.pack(anchor="w", pady=5)
                checkboxes.append(checkbox)

        # Use Toplevel instead of Tk to create a popup window
        window = tk.Toplevel()
        window.title("Scrollable Checkboxes")
        window.geometry("400x400")

        # Create a canvas and scrollbar
        canvas = tk.Canvas(window)
        scrollbar = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Configure canvas to work with scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configure canvas to work with mouse
        canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        # Create a frame inside the canvas
        scrollable_frame = tk.Frame(canvas)

        # Create a window inside the canvas that contains the frame
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Bind the frame size changes to update the canvas scrollregion
        scrollable_frame.bind("<Configure>", on_frame_configure)

        item_list = list(self.controller.investors)

        # Lists to store checkbox variables and checkbox widgets
        vars = []
        checkboxes = []

        # Create checkboxes for each item in the list
        create_checkboxes()

        window.mainloop()
     



class OutputPage(tk.Frame):
    def __init__(self, parent, controller):
        def resize(event):
            # Set the sash position to half of the current window width
            paned_window.sash_place(0, self.winfo_width() // 2, 1)


        def toggle_bulk_entry():
            # Show or hide the entry widget based on the selected option
            if controller.bulk_choice.get():
                self.bulk_frame.pack()
            else:
                self.bulk_frame.pack_forget()

        def create_sample():
            controller.is_sample = True
        
            # Define the path to the sample folder
            folder_path = 'sample/'

            # Use glob to find the file (assuming any file with any extension)
            file_path = glob.glob(os.path.join(folder_path, '*'))  # Matches any file
            
            if (len(file_path) == 1):
                os.remove(file_path[0])

            controller.submit_action()

            # Use glob to find the file (assuming any file with any extension)
            file_path = glob.glob(os.path.join(folder_path, '*'))  # Matches any file

            # Ensure there is exactly one file
            if len(file_path) == 1:
                print(f"The file path is: {file_path[0]}")
            else:
                print("Error: Either no files or multiple files found in the folder.")
                return

            sample_output(controller.frames["OutputPage"].frame_sample, file_path[0])

            controller.is_sample = False
            

        super().__init__(parent)
        self.controller = controller

        # Create a PanedWindow to split the window
        paned_window = tk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # Create the left side with a sidebar and content area
        content_frame = tk.Frame(paned_window, bg="lightblue")
        paned_window.add(content_frame)

        # Left content area (where the content will change)
        left_frame = tk.Frame(content_frame, bg="lightblue")
        left_frame.pack(fill=tk.BOTH, expand=True)

        # Create a frame for directory selection
        frame_dir = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_dir.pack(padx=20, pady=20, fill="x")

        label_dir_title = tk.Label(frame_dir, text="Select Where To Store Output PDF", font=controller.font_label, bg="#e6e6e6")
        label_dir_title.pack(anchor="w")

        button_dir = tk.Button(frame_dir, text="Select Output Directory", command=controller.select_directory, font=controller.font_button, bg="#4CAF50", fg="white")
        button_dir.pack(side="left", padx=10, pady=5)

        controller.label_dir = tk.Label(frame_dir, textvariable=controller.selected_directory, bg="#e6e6e6", font=controller.font_label)
        controller.label_dir.pack(side="left", padx=10)

        # Create a frame for choosing between split and bulk
        frame_output_choice = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_output_choice.pack(padx=20, pady=20, fill="x")

        label_options = tk.Label(frame_output_choice, text="Select Output Format", font=controller.font_label, bg="#e6e6e6")
        label_options.pack(anchor="w")

        split_option = tk.Checkbutton(frame_output_choice, text="Split", variable=controller.split_choice, command=toggle_bulk_entry)
        bulk_option = tk.Checkbutton(frame_output_choice, text="Bulk", variable=controller.bulk_choice, command=toggle_bulk_entry)

        split_option.pack(anchor="w")
        bulk_option.pack(anchor="w")

        frame_delimiters = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_delimiters.pack(padx=20, pady=20, fill="x")

        tk.Label(frame_delimiters, text="Start Delimiter:").grid(row=1, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.start_delimiter).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_delimiters, text="End Delimiter:").grid(row=2, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.end_delimiter).grid(row=2, column=1, padx=5, pady=5)

        # Create an Entry widget for bulk input, but don't pack it initially
        self.bulk_frame = tk.Frame(frame_output_choice)

        tk.Label(self.bulk_frame, text="Output File Name:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        tk.Entry(self.bulk_frame, textvariable=controller.output_file).grid(row=3, column=1, padx=5, pady=5)

        # Right side for displaying an image
        self.frame_sample = tk.Frame(paned_window, bg="lightgreen")
        sample_button = tk.Button(self.frame_sample, text="Create Sample", font=("Arial", 12), command=create_sample)
        sample_button.pack()

        paned_window.add(self.frame_sample)

        # Bottom frame for navigation buttons
        bottom_frame = tk.Frame(self, bg="lightgray", height=50)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Add Back and Next buttons, and center them
        back_button = tk.Button(bottom_frame, text="Back", font=("Arial", 12), command=lambda: controller.show_frame("InputPage"))
        generate_button = tk.Button(bottom_frame, text="Generate", font=("Arial", 12), command = controller.submit_action)

        # Use a frame to center both buttons in the middle of the bottom frame
        button_frame = tk.Frame(bottom_frame, bg="lightgray")
        button_frame.pack(expand=True)

        # Pack buttons inside the center frame
        back_button.pack(side=tk.LEFT, padx=10, pady=10)
        generate_button.pack(side=tk.LEFT, padx=10, pady=10)

        # Bind the window resize event to the resize function
        self.bind("<Configure>", resize)

        """
        label = tk.Label(self, text="GPES FileGen", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)
        """

        """
        frame_left = tk.Frame(self)
        frame_left.pack(side = "left", fill = "both")

        # Create a frame for directory selection
        frame_dir = tk.Frame(frame_left, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_dir.pack(padx=20, pady=20, fill="x")

        label_dir_title = tk.Label(frame_dir, text="Select Where To Store Output PDF", font=controller.font_label, bg="#e6e6e6")
        label_dir_title.pack(anchor="w")

        button_dir = tk.Button(frame_dir, text="Select Output Directory", command=controller.select_directory, font=controller.font_button, bg="#4CAF50", fg="white")
        button_dir.pack(side="left", padx=10, pady=5)

        controller.label_dir = tk.Label(frame_dir, textvariable=controller.selected_directory, bg="#e6e6e6", font=controller.font_label)
        controller.label_dir.pack(side="left", padx=10)

        # Create a frame for choosing between split and bulk
        frame_output_choice = tk.Frame(frame_left, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_output_choice.pack(padx=20, pady=20, fill="x")

        label_options = tk.Label(frame_output_choice, text="Select Output Format", font=controller.font_label, bg="#e6e6e6")
        label_options.pack(anchor="w")

        split_option = tk.Checkbutton(frame_output_choice, text="Split", variable=controller.split_choice, command=toggle_bulk_entry)
        bulk_option = tk.Checkbutton(frame_output_choice, text="Bulk", variable=controller.bulk_choice, command=toggle_bulk_entry)

        split_option.pack(anchor="w")
        bulk_option.pack(anchor="w")

        frame_delimiters = tk.Frame(frame_left, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_delimiters.pack(padx=20, pady=20, fill="x")

        tk.Label(frame_delimiters, text="Start Delimiter:").grid(row=1, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.start_delimiter).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_delimiters, text="End Delimiter:").grid(row=2, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.end_delimiter).grid(row=2, column=1, padx=5, pady=5)

        # Create an Entry widget for bulk input, but don't pack it initially
        self.bulk_frame = tk.Frame(frame_output_choice)

        tk.Label(self.bulk_frame, text="Output File Name:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        tk.Entry(self.bulk_frame, textvariable=controller.output_file).grid(row=3, column=1, padx=5, pady=5)


        

        # Create a frame to show sample pdf
        self.frame_sample = Frame(self)
        self.frame_sample.pack(side = "left", fill = "both")

        

        frame_buttons = tk.Frame(self)
        frame_buttons.pack(side = "bottom", fill = "both")
        # Button to go back a page
        back_button = tk.Button(frame_buttons, text="Back",
                                command=lambda: controller.show_frame("InputPage"))
        back_button.pack()


        submit_button = tk.Button(frame_buttons, text="Submit", command = controller.submit_action)
        submit_button.pack()

        """




if __name__ == "__main__":
    # Define the path to the sample folder
    folder_path = 'sample/'
    # Use glob to find the file (assuming any file with any extension)
    file_path = glob.glob(os.path.join(folder_path, '*'))  # Matches any fil
    # Ensure there is exactly one file
    if len(file_path) == 1:
        print(f"The file path is: {file_path[0]}")
    else:
        print("Error: Either no files or multiple files found in the folder.")
        
    if len(file_path) == 1:
        os.remove(file_path[0])

    app = MainApp()
    app.mainloop()
