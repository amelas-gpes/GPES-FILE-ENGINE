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

from parse_input_excel import *

from pdf_viewer import *

class MainApp(tk.Tk):
    def select_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.label_file.config(text=f"Selected: {file_path}")
            self.selected_file.set(file_path)

            parse_input_excel(file_path, "Allocation", self.fund_info, self.inv_info, self.total_fund_info)

            for inv_num in self.inv_info:
                self.investors.add(self.inv_info[inv_num]["Partner Name"])
            self.checked_investors = list(self.investors)

            self.frames["InputPage"].label_confirmation.config(text=f"{len(self.inv_info)} Investors, {len(self.inv_info[next(iter(self.inv_info))])} Fields")

            #Populate file list
            #self.field_list = ["<Investment #1>"]
            inv_num = next(iter(self.inv_info))
            for key in self.inv_info[inv_num]:
                self.field_list.append("<" + key + ">")
            
            for index, val in enumerate(self.cap_call_table_data):
                self.frames["OutputPage"].create_entry_pair(val, index)
            

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

        self.fund_info["Contact Name"] = first_name + " " + last_name
        self.fund_info["Contact Email"] = email
        
        #go through each investors and create pdfs for them
        for inv in self.inv_info:
            if self.inv_info[inv]["Partner Name"] not in self.checked_investors:
                continue
        
            if option == "Capital Call":
                doc = Document(r".\documents\cap_call_template.docx")
                
                
                output_pdf_name = self.inv_info[inv]["Partner Name"]
                if self.is_sample:
                    output_pdf_name = "sample"
                if self.output_file_split.get() != "":
                    if (".pdf" not in self.output_file_split.get()):
                        self.output_file_split.set(self.output_file_split.get() + ".pdf")
                    output_pdf_name = self.output_file_split.get()

                    output_pdf_name.replace("<investor_name>", self.inv_info[inv]["Parner Name"])
                    print("OUTPUT_PDF_NAME: ", output_pdf_name)

                output_path = os.path.join(output_directory, f"{output_pdf_name}")
                if not self.is_sample:
                    self.files_list.append(output_path+".pdf")
                

                create_cap_call_pdf(doc, excel_file_path, self.fund_info, self.inv_info[inv], self.total_fund_info, output_path, logo_path, self.text_fields, start_delim, end_delim, self.cap_call_table_data)

            if self.is_sample:
                self.sample_ready = True
                break

        if (self.bulk_choice.get()):
            self.merge_pdfs_in_folder(output_directory, self.output_file.get())

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

    def on_selected_option_change(self, *args):
        # Trigger the child to update its frame based on var
        self.frames["OutputPage"].update_frame_content()

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
        self.selected_option = tk.StringVar(self)
        self.selected_option.trace_add("write", self.on_selected_option_change)

        self.first_name = tk.StringVar(self, value="")
        self.last_name = tk.StringVar(self, value="")
        self.email = tk.StringVar(self, value="")

        self.label_file = ""
        self.label_logo = ""
        self.label_dir = ""

        self.investors = set()
        self.checked_investors = list(self.investors)

        self.is_sample = False

        self.inv_info = dict()
        self.fund_info = dict()
        self.total_fund_info = dict()

        self.confirmation_data = "0 fields, 0 investors"

        self.cap_call_table_data = [
            ["Investment", "<Investment #1>", "<Investment #1>"],
            ["Management Fees", "<Gross Mgmt Fee>", "<Gross Mgmt Fee>"],
            ["Partnership Expenses", "<Pshp Exp>", "<Pshp Exp>"]
        ]
        """Modifiable table entries"""
        self.row_count = 3  # Keeps track of the number of rows

        # Create a StringVar to hold the choice (split or bulk)
        self.split_choice = tk.IntVar(value=1)
        self.bulk_choice = tk.IntVar(value=0)
        self.start_delimiter = tk.StringVar(self, value = "<start>")
        self.end_delimiter = tk.StringVar(self, value = "<end>")
        self.output_file = tk.StringVar(self, value = "bulk.pdf")
        self.output_file_split = tk.StringVar(self, value="")

        #text fields for capital call:
        self.cap_call_field1 = tk.StringVar(value="""In accordance with Section <Section#> of the Amended and Restated Limited Partnership Agreement of the Fund dated <Notice_Date> (the “Agreement”), the Fund is calling capital for Investments, Management Fees, and Partnership Expenses. Capitalized terms used but not defined in this notice are defined in the Agreement.""")
        self.cap_call_field2 = tk.StringVar(value="""Additional detail on this capital notice is presented on the attached Exhibit A.""")
        self.cap_call_field3 = tk.StringVar(value="""Your portion of the call is <Total_Amount_Due> and is due on <Due_Date>.  Please send your payment by wire transfer in accordance with the instructions provided below.""")
        self.cap_call_field4 = tk.StringVar(value="""Should you have any questions on this notice, please do not hesitate to contact <Contact_Name> at <Contact_Email>.""")
        
        self.cap_call_text = self.cap_call_field1.get() + "\n\n" + self.cap_call_field2.get() + "\n\n" + self.cap_call_field3.get() + "\n\n" + self.cap_call_field4.get()
        self.text_fields = self.cap_call_text.split("\n\n")

        self.files_list = []

        self.field_list = []

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
        super().__init__(parent)
        self.controller = controller


        label = tk.Label(self, text="GPES Mail Merge Automation", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)

        # Frame for file selection
        frame_file = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_file.pack(padx=20, pady=20, fill="x")

        label_file_title = tk.Label(frame_file, text="Select Allocation File (Please ensure Investor Number is first column)", font=controller.font_label, bg="#e6e6e6")
        label_file_title.pack(anchor="w")

        button_file = tk.Button(frame_file, text="Select File", command=controller.select_file, font=controller.font_button, bg="#4A164B", fg="white")
        button_file.pack(side="left", padx=10, pady=5)

        controller.label_file = tk.Label(frame_file, textvariable=controller.selected_file, bg="#e6e6e6", font=controller.font_label)
        controller.label_file.pack(side="left", padx=10)

        # Frame for confirmation modal
        frame_confirmation = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_confirmation.pack(padx=20, pady=20, fill="x")

        self.label_confirmation = tk.Label(frame_confirmation, text=controller.confirmation_data, font=controller.font_label, bg="#e6e6e6")
        self.label_confirmation.pack(anchor="w")

        # Frame for logo selection
        frame_logo = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_logo.pack(padx=20, pady=20, fill="x")

        label_logo_title = tk.Label(frame_logo, text="Select Firm Logo", font=controller.font_label, bg="#e6e6e6")
        label_logo_title.pack(anchor="w")

        button_logo = tk.Button(frame_logo, text="Select Logo", command=controller.select_logo, font=controller.font_button, bg="#4A164B", fg="white")
        button_logo.pack(side="left", padx=10, pady=5)

        controller.label_logo = tk.Label(frame_logo, textvariable=controller.selected_logo, bg="#e6e6e6", font=controller.font_label)
        controller.label_logo.pack(side="left", padx=10)

        # Frame for the option menu
        frame_option = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_option.pack(padx=20, pady=20, fill="x")

        label_option_title = tk.Label(frame_option, text="Select Document Type", font=controller.font_label, bg="#e6e6e6")
        label_option_title.pack(anchor="w")

        options = ["Capital Call", "Distribution Notice"]
        
        #controller.selected_option.set(options[0])

        option_menu = tk.OptionMenu(frame_option, controller.selected_option, *options, command=controller.option_selected)
        option_menu.config(font=controller.font_button, bg="#4A164B", fg="white")
        option_menu.pack(side="left", padx=10, pady=5)

        controller.label_option = tk.Label(frame_option, text="No option selected", bg="#e6e6e6", font=controller.font_label)
        controller.label_option.pack(side="left", padx=10)


        # Frame for contact info
        frame_contacts = tk.Frame(self, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_contacts.pack(padx=20, pady=20, fill="x")

        label_contacts = tk.Label(frame_contacts, text="Contact Information - Add Contact Information to be Applied to Each Document", font=controller.font_label, bg="#e6e6e6")
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
        
        # Bottom frame for navigation buttons
        bottom_frame = tk.Frame(self, bg="lightgray", height=50)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Add Back and Next buttons, and center them
        next_button = tk.Button(bottom_frame, text="Next", font=("Arial", 12), command=lambda: controller.show_frame("OutputPage"))

        # Use a frame to center both buttons in the middle of the bottom frame
        button_frame = tk.Frame(bottom_frame, bg="lightgray")
        button_frame.pack(expand=True)

        # Pack buttons inside the center frame
        next_button.pack(side=tk.LEFT, padx=10, pady=10)

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

        def toggle_split_entry():
            if controller.split_choice.get():
                self.split_frame.pack()
            else:
                self.split_frame.pack_forget()

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

        # Create the left side with a content area (which will be scrollable)
        content_frame = tk.Frame(paned_window, bg="#e6e6e6")
        paned_window.add(content_frame)

        # Create a canvas and scrollbar for the left frame
        canvas = tk.Canvas(content_frame, bg="#e6e6e6")
        scrollbar = tk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Place the canvas and scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a frame inside the canvas
        self.scrollable_frame = tk.Frame(canvas, bg="#e6e6e6")

        # This is needed to allow the scrollable frame to adjust to the canvas size
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Add the scrollable frame to the canvas
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # Left content area (where the content will change)
        left_frame = tk.Frame(self.scrollable_frame, bg="#e6e6e6")
        left_frame.pack(fill=tk.BOTH, expand=True)

        # Bind mouse wheel scrolling
        canvas.bind_all("<MouseWheel>", lambda event: self._on_mouse_wheel(event, canvas))

        # Create a frame for directory selection
        frame_dir = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_dir.pack(padx=20, pady=20, fill="x")

        label_dir_title = tk.Label(frame_dir, text="Select Where To Store Output PDF", font=controller.font_label, bg="#e6e6e6")
        label_dir_title.pack(anchor="w")

        button_dir = tk.Button(frame_dir, text="Select Output Directory", command=controller.select_directory, font=controller.font_button, bg="#4A164B", fg="white")
        button_dir.pack(side="left", padx=10, pady=5)

        controller.label_dir = tk.Label(frame_dir, textvariable=controller.selected_directory, bg="#e6e6e6", font=controller.font_label)
        controller.label_dir.pack(side="left", padx=10)

        # Create a frame for choosing between split and bulk
        frame_output_choice = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_output_choice.pack(padx=20, pady=20, fill="x")

        label_options = tk.Label(frame_output_choice, text="Select Output Format", font=controller.font_label, bg="#e6e6e6")
        label_options.pack(anchor="w")

        split_option = tk.Checkbutton(frame_output_choice, text="Split", variable=controller.split_choice, command=toggle_split_entry)
        bulk_option = tk.Checkbutton(frame_output_choice, text="Bulk", variable=controller.bulk_choice, command=toggle_bulk_entry)

        split_option.pack(anchor="w")
        bulk_option.pack(anchor="w")

        frame_delimiters = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        frame_delimiters.pack(padx=20, pady=20, fill="x")

        tk.Label(frame_delimiters, text="Start Delimiter:").grid(row=1, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.start_delimiter).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame_delimiters, text="End Delimiter:").grid(row=2, column=0, padx=5, pady=5, sticky='e')

        tk.Entry(frame_delimiters, textvariable=controller.end_delimiter).grid(row=2, column=1, padx=5, pady=5)

        # Create an entry widget for split input output file name, dont pack initially
        self.split_frame = tk.Frame(frame_output_choice)
        tk.Label(self.split_frame, text="Output File Name for Split:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        tk.Entry(self.split_frame, textvariable=controller.output_file_split).grid(row=3, column=1, padx=5, pady=5)
        tk.Label(self.split_frame, text="""Placeholders:\n<investor_name> : investor name""").grid(row=4, column=0, padx=5, pady=5, sticky="e")

        toggle_split_entry()

        # Create an Entry widget for bulk input, but don't pack it initially
        self.bulk_frame = tk.Frame(frame_output_choice)

        tk.Label(self.bulk_frame, text="Output File Name for Bulk:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        tk.Entry(self.bulk_frame, textvariable=controller.output_file).grid(row=3, column=1, padx=5, pady=5)

        #Modifiable entries
        # Frame for contact info
        self.content_frame = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        self.content_frame.pack(padx=20, pady=20, fill="x")

        # Frame for file selection
        self.table_frame = tk.Frame(left_frame, bg="#e6e6e6", bd=2, relief="sunken", padx=10, pady=10)
        self.table_frame.pack(padx=20, pady=20, fill="x")

        # Use grid layout inside the window_frame
        self.table_frame.grid_columnconfigure(1, weight=1)  # Make second column expandable

        # Add Button to dynamically add rows
        add_button = tk.Button(left_frame, text="Add New Entry Pair", command=self.add_entry_pair)
        add_button.pack(pady=10)


        # Right side for displaying an image with a top frame for the button
        self.frame_right = tk.Frame(paned_window, bg="#F1EAE5")  # The main right-side frame
        self.frame_top = tk.Frame(self.frame_right, bg="#F27D42")  # Top frame for the button
        self.frame_sample = tk.Frame(self.frame_right, bg="#F1EAE5")  # Frame below for content display

        # Create and pack the button into the top frame
        sample_button = tk.Button(self.frame_top, text="Create Sample", font=("Arial", 12), command=create_sample)
        sample_button.pack(pady=10)  # Add padding to space it out a bit

        # Pack the frames vertically
        self.frame_top.pack(fill="x")  # Top frame with button
        self.frame_sample.pack(expand=True, fill="both")  # Content frame will expand to take remaining space

        # Add the main right-side frame to the paned window
        paned_window.add(self.frame_right)

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

    def _on_mouse_wheel(self, event, canvas):
        """Scroll the canvas vertically when the mouse wheel is used (Windows)."""
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def update_frame_content(self):
        var_value = self.controller.selected_option.get()

        # Clear the frame by destroying all existing widgets
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        # Add new widgets based on the value of the variable
        if var_value == "Capital Call":
            top_frame = tk.Frame(self.content_frame)
            top_frame.pack(side="top", fill="x", pady=10)

            tk.Label(top_frame, text="Customize Capital Call").pack(side="left", padx=10, pady=10)

            tk.Button(top_frame, text="Reset", command=self.onReset).pack(side="right", padx=10, pady=10)

            self.cap_call_text_field = tk.Text(self.content_frame, wrap="word", width=100, undo=True)
            self.cap_call_text_field.pack(side="bottom", pady=10)

            self.cap_call_text_field.insert("1.0", self.controller.cap_call_text)
            self.cap_call_text_field.bind("<KeyRelease>", self.update_cap_call)

        elif var_value == "Distribution Notice":
            tk.Label(self.content_frame, text="Customize Distributed Notice").pack(pady=10)
            text = tk.Text(self.content_frame)
            text.pack(pady=10)

    
        
    #Updates cap call modifiable text field
    def update_cap_call(self, event=None):
        self.controller.text_fields = self.cap_call_text_field.get("1.0", "end-1c").split("\n\n")
    
    #Resets cap call modifiable text field
    def onReset(self):
        self.controller.text_fields = self.controller.cap_call_text.split("\n\n")
        self.cap_call_text_field.delete('1.0', tk.END)
        self.cap_call_text_field.insert("1.0", self.controller.cap_call_text)

    #Creates entries for modifiable table
    def create_entry_pair(self, default, row):
        """Helper function to create a pair of left and right entries in a given row."""
        left_entry = tk.Entry(self.table_frame)
        left_entry.insert(0, default[0])
        left_entry.grid(row=row, column=0, padx=5, pady=5, sticky="w")

        left_entry.bind("<KeyRelease>", lambda event, r=row: self.update_list(r, 0, left_entry.get()))


        first_entry_val = tk.StringVar(self.table_frame)
        first_entry_val.set("Select an option")
        first_entry_val.trace("w", lambda *args, r=row: self.update_list(r, 1, first_entry_val.get()))

        first_entry = tk.OptionMenu(self.table_frame, first_entry_val, *(self.controller.field_list))
        first_entry.grid(row=row, column=1, padx=5, pady=5, sticky="e")

        
        second_entry_val = tk.StringVar(self.table_frame)
        second_entry_val.set("Select an option")
        second_entry_val.trace("w", lambda *args, r=row: self.update_list(r, 2, second_entry_val.get()))

        second_entry = tk.OptionMenu(self.table_frame, second_entry_val, *(self.controller.field_list))
        second_entry.grid(row=row, column=2, padx=5, pady=5, sticky="e")

    def update_list(self, row, column, new_value):
        """Update the list based on the entry changes."""
        self.controller.cap_call_table_data[row][column] = new_value
        print(self.controller.cap_call_table_data)

    def add_entry_pair(self):
        """Function to add a new pair of entries dynamically."""
        self.create_entry_pair(["New Entry", "", ""], self.controller.row_count)
        self.controller.cap_call_table_data.append(["New Entry", "", ""])
        self.controller.row_count += 1  # Increment the row count each time a new pair is added

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