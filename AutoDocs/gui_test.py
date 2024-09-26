import tkinter as tk

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Page Switcher")

        # Create a container to hold the frames (pages)
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Dictionary to hold pages
        self.frames = {}

        # Initialize all the pages
        for F in (HomePage, PageOne, PageTwo):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Show the home page initially
        self.show_frame("HomePage")

    def show_frame(self, page_name):
        # Bring the frame to the front
        frame = self.frames[page_name]
        frame.tkraise()


class HomePage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = tk.Label(self, text="Home Page", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)

        # Button to go to PageOne
        button1 = tk.Button(self, text="Next",
                            command=lambda: controller.show_frame("PageOne"))
        button1.pack()


class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = tk.Label(self, text="This is Page One", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)

        # Button to go back to the home page
        home_button = tk.Button(self, text="Back",
                                command=lambda: controller.show_frame("HomePage"))
        home_button.pack()

        next_button = tk.Button(self, text="Next", command=lambda: controller.show_frame("PageTwo"))
        next_button.pack()


class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = tk.Label(self, text="This is Page Two", font=('Arial', 16))
        label.pack(side="top", fill="x", pady=10)

        # Button to go back to the home page
        home_button = tk.Button(self, text="Back",
                                command=lambda: controller.show_frame("PageOne"))
        home_button.pack()


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
