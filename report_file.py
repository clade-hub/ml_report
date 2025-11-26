import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import fitz  # PyMuPDF
import os
from odf.opendocument import OpenDocumentText
from odf.draw import Frame, Image
from odf.text import P
from datetime import datetime

def create_report(patient_title, patient_name, patient_dob, image_path, image_width_px, image_height_px, odt_path):
    doc = OpenDocumentText()
    
    # Headline
    today = datetime.now().strftime("%d.%m.%Y")
    headline = P(text=f"Befundbericht zur Bewegungsanalyse vom {today}")
    doc.text.addElement(headline)
    
    # Patient Info
    patient_info1 = P(text=f"{patient_title} {patient_name}")
    patient_info2 = P(text=f"geboren am: {patient_dob}")
    doc.text.addElement(patient_info1)
    doc.text.addElement(patient_info2)

    # --- Aspect Ratio Calculation ---
    # Define a max width for the image in the document (e.g., 16cm)
    max_width_cm = 16.0
    
    # Calculate aspect ratio to avoid distortion
    if image_width_px > 0:
        aspect_ratio = image_height_px / image_width_px
    else:
        aspect_ratio = 1 # Avoid division by zero

    # Calculate the new height to maintain the aspect ratio
    frame_width_cm = max_width_cm
    frame_height_cm = max_width_cm * aspect_ratio

    # Add image with correct aspect ratio
    frame = Frame(width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", x="2.5cm", y="4.5cm")
    href = doc.addPicture(image_path)
    frame.addElement(Image(href=href))
    doc.text.addElement(frame)

    doc.save(odt_path)

from PIL import Image as PILImage, ImageTk

class CoordinatePicker:
    def __init__(self, parent, page_pixmap):
        self.parent = parent
        self.page_pixmap = page_pixmap
        self.coords = []
        self.top = tk.Toplevel(parent)
        self.top.title("Click to select snippet area")

        # Convert pixmap to a PhotoImage
        img = PILImage.frombytes("RGB", [page_pixmap.width, page_pixmap.height], page_pixmap.samples)
        self.photo_image = ImageTk.PhotoImage(image=img)

        self.canvas = tk.Canvas(self.top, width=page_pixmap.width, height=page_pixmap.height)
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo_image)
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.on_click)

    def on_click(self, event):
        self.coords.append(event.x)
        self.coords.append(event.y)
        
        if len(self.coords) == 2:
            print(f"Top-left corner selected at ({event.x}, {event.y})")
        elif len(self.coords) == 4:
            print(f"Bottom-right corner selected at ({event.x}, {event.y})")
            self.top.destroy()

    def get_coords(self):
        self.parent.wait_window(self.top)
        return self.coords if len(self.coords) == 4 else None

def generate_report():
    patient_title = simpledialog.askstring("Input", "Anrede (Herr/Frau):")
    if not patient_title: return
    
    patient_name = simpledialog.askstring("Input", "Patient Name:")
    if not patient_name: return

    patient_dob = simpledialog.askstring("Input", "Geboren am (DD.MM.YYYY):")
    if not patient_dob: return

    pdf_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return

    try:
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        
        page_num = simpledialog.askinteger(
            "Select Page Number",
            f"Enter page number (1 to {num_pages}):",
            minvalue=1,
            maxvalue=num_pages
        )
        if page_num is None: return

        # Interactive coordinate selection
        page = doc.load_page(page_num - 1)
        pix = page.get_pixmap()
        
        root.withdraw() # Hide the main window
        picker = CoordinatePicker(root, pix)
        coords = picker.get_coords()
        root.deiconify() # Show the main window again

        if not coords:
            messagebox.showinfo("Cancelled", "Coordinate selection was cancelled.")
            return
        
        rect = fitz.Rect(coords)
        pix = page.get_pixmap(clip=rect)
        
        image_path = "temp_pdf_snippet.png"
        pix.save(image_path)

        # Get snippet dimensions for aspect ratio calculation
        with PILImage.open(image_path) as img:
            snippet_width, snippet_height = img.size
        
        doc.close()

        odt_path = filedialog.asksaveasfilename(
            title="Save ODT report as",
            defaultextension=".odt",
            filetypes=[("ODT files", "*.odt")]
        )
        if not odt_path:
            os.remove(image_path)
            return

        create_report(patient_title, patient_name, patient_dob, image_path, snippet_width, snippet_height, odt_path)
        os.remove(image_path)
        
        messagebox.showinfo("Success", f"Successfully created {odt_path}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Create the main window
root = tk.Tk()
root.title("Report Generator")

# Create and place the widgets
main_label = tk.Label(root, text="Generate a new report.")
main_label.pack(pady=10)

generate_button = tk.Button(root, text="Generate Report", command=generate_report)
generate_button.pack(pady=20)

# Start the main event loop
root.mainloop()

print("Report generator closed. Goodbye, Carlo!")

