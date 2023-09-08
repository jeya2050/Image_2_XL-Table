import tkinter as tk
from tkinter import *
from tkinter import filedialog
from nanonets import NANONETSOCR
from PIL import ImageTk, Image  

class Image2xl:
    
    def __init__(self, root):
        self.canvas = tk.Canvas(root, width = 500,height = 250)  
        self.canvas.pack() 
        self.model = NANONETSOCR()
        self.model.set_token('89eeebff-4d5f-11ee-a899-62c99726f28c')
        # Initialize the main application window
        self.root = root
        self.root.title("PDF to Excell")

        file_frame = tk.LabelFrame(root, text="file upload")
        file_frame.place(height=400, width=550, x=3, y=5)
            
        # Create an input frame for organizing UI elements
        self.input_frame = tk.Frame(root, width=300, height=200)
        self.input_frame.pack(side="top", padx=10, pady=10)

        # Button to upload an Excel file
        self.upload_button = tk.Button(
            file_frame, text="Upload IMAGE File", command=self.upload_file)
        self.upload_button.pack(pady=10)
        self.upload_button.place(x=10, y=10)

        # Button to display selected data
        self.display_button = tk.Button(
            file_frame, text="CONVERT to EXCEL", command=self.conv_excel)
        self.display_button.pack(pady=5)
        self.display_button.place(x=10, y=150)

          # Reset Button
        self.reset_button = tk.Button(
            file_frame, text="Reset", command=self.reset_ui)
        self.reset_button.pack(pady=5)
        self.reset_button.place(x=200, y=150)  # Adjust position as needed

    def conv_excel(self):
        self.model.convert_to_csv(f'{self.file_path}',output_file_name="result.csv")
        self.l= tk.Label(root, text = "CONVERSION DONE")
        self.l.config(font =("Courier", 14))
        self.l.place(x=100, y=250)

    def upload_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("images", "*.png")])
        if self.file_path:
            width=570
            height=700
            image = Image.open(self.file_path)
            self.resize_image = image.resize((width, height))
            self.photo_image = ImageTk.PhotoImage(self.resize_image)

            # Create a label and put the image on it
            self.label = tk.Label(root,image=self.photo_image)
            self.label.place(x=600, y=20)
            self.l3 = tk.Label(root, text = "IMAGE UPLOADED")
            self.l3.config(font =("Courier", 14))
            self.l3.place(x=150, y=30)

    def reset_ui(self):
        self.l.destroy()
        self.l3.destroy()
        self.label.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = Image2xl(root)
    root.mainloop()
