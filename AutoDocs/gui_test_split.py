import tkinter as tk

def toggle_bulk_entry():
    # Show or hide the entry widget based on the selected option
    if choice.get() == "bulk":
        bulk_entry.pack(pady=5)
    else:
        bulk_entry.pack_forget()

# Create the main window
root = tk.Tk()
root.title("Split or Bulk Choice")

# Create a StringVar to hold the choice (split or bulk)
choice = tk.StringVar(value="split")

# Create and pack the radiobuttons for "Split" and "Bulk"
split_option = tk.Radiobutton(root, text="Split", variable=choice, value="split", command=toggle_bulk_entry)
bulk_option = tk.Radiobutton(root, text="Bulk", variable=choice, value="bulk", command=toggle_bulk_entry)

split_option.pack(anchor="w")
bulk_option.pack(anchor="w")

# Create an Entry widget for bulk input, but don't pack it initially
bulk_entry = tk.Entry(root)

# Start the Tkinter event loop
root.mainloop()
