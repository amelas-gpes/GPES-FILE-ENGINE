import fitz
from tkinter import *
from PIL import Image, ImageTk

def sample_output(root, doc):
    # transformation matrix we can apply on pages
    zoom = 1
    mat = fitz.Matrix(zoom, zoom)

    # count number of pages
    num_pages = len(doc)

    # add scroll bar
    scrollbar = Scrollbar(root)
    scrollbar.pack(side = RIGHT, fill = Y)

    # add canvas
    canvas = Canvas(root, yscrollcommand = scrollbar.set)
    canvas.pack(side = LEFT, fill = BOTH, expand = 1)

    # define entry point (field for taking inputs)
    entry = Entry(root)

    # add a label for the entry point
    label = Label(root, text="Enter page number to display:")

    # persistent image variable to prevent garbage collection
    image_reference = None

    def pdf_to_img(page_num):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=mat)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    def show_image():
        nonlocal image_reference
        try:
            page_num = int(entry.get()) - 1

            if page_num < 0 or page_num >= num_pages:
                print(f"Page number out of bounds: {page_num + 1}")
                return

            im = pdf_to_img(page_num)
            img_tk = ImageTk.PhotoImage(im)

            # Create a frame to hold the image
            frame = Frame(canvas)
            panel = Label(frame, image=img_tk)
            panel.pack(side="bottom", fill="both", expand="yes")

            # Update the canvas with the new image
            canvas.create_window(0, 0, anchor='nw', window=frame)

            # Set canvas size to image size
            canvas.config(scrollregion=canvas.bbox("all"))
            #canvas.config(width=img_tk.width(), height=img_tk.height())

            # Keep a reference to the image to prevent garbage collection
            image_reference = img_tk

        except Exception as e:
            print(f"Error displaying page: {e}")

    # add button to display pages
    button = Button(root, text="Show Page", command=show_image)

    # set visual locations
    label.pack(side=TOP, fill=None)
    entry.pack(side=TOP, fill=BOTH)
    button.pack(side=TOP, fill=None)

    entry.insert(0, '1')
    show_image()  # Display the first page initially

    scrollbar.config(command=canvas.yview)


if __name__ == "__main__":
    # initialize and set screen size
    root = Tk()
    root.geometry('750x700')

    # open pdf file
    file_name = r"C:\Users\ppark\OneDrive - GP Fund Solutions, LLC\Desktop\GPES-FILE-ENGINE\AutoDocs\output\quarterly_updates\bulk.pdf"
    doc = fitz.open(file_name)


    # Create a frame to show sample pdf
    frame_sample = Frame(root)
    frame_sample.pack()

    sample_output(frame_sample, doc)

    root.mainloop()
    doc.close()
