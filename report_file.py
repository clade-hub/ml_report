import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import fitz  # PyMuPDF
import os
from odf.opendocument import OpenDocumentText
from odf.draw import Frame, Image
from odf.text import P
from odf.table import Table, TableColumn, TableRow, TableCell # Import table elements
from datetime import datetime

from PIL import Image as PILImage, ImageTk # Ensure PIL is imported here for image processing

def create_report(patient_full_title, patient_name, patient_dob, image_path, image_width_px, image_height_px, odt_path, logo_path=None):
    doc = OpenDocumentText()
    
    # Create a table for the header layout (logo and text side-by-side)
    header_table = Table(name="HeaderLayoutTable")
    
    # Define columns for the table
    # These widths are for layout. Total content width for A4 is usually around 18cm.
    # Temporarily removed width specification to resolve error.
    # Correct way to set column width in odfpy needs further investigation (likely via styles or TableColumnProperties).
    logo_col_width_cm = 6.0 # Column for the logo (desired width)
    text_col_width_cm = 12.0 # Column for the text info (desired width)
    
    header_table.addElement(TableColumn()) # Removed columnwidth attribute
    header_table.addElement(TableColumn()) # Removed columnwidth attribute

    header_row = TableRow()

    # --- Logo Cell (Left Side) ---
    logo_cell = TableCell()
    if logo_path:
        try:
            with PILImage.open(logo_path) as logo_img:
                original_logo_width_px, original_logo_height_px = logo_img.size

            max_logo_width_cm = 5.0 # Increased logo size
            
            if original_logo_width_px > 0:
                aspect_ratio = original_logo_height_px / original_logo_width_px
            else:
                aspect_ratio = 1 # Fallback for invalid width

            # Calculate the new height in cm to maintain aspect ratio
            # Ensure the logo fits within its column, with some padding
            # We subtract a small margin (e.g., 0.5cm) from the column width
            # to prevent the image from touching the cell edges.
            # For now, we'll use the max_logo_width_cm as the target width for the frame
            # as the column itself is auto-sizing.
            logo_frame_width_cm = max_logo_width_cm
            logo_frame_height_cm = logo_frame_width_cm * aspect_ratio

            # Create a paragraph to hold the logo frame within the cell
            logo_p = P()
            logo_frame = Frame(
                name="LogoFrame",
                width=f"{logo_frame_width_cm}cm",
                height=f"{logo_frame_height_cm}cm",
                anchortype="paragraph" # Anchor to the paragraph within the cell
            )
            logo_href = doc.addPicture(logo_path)
            if logo_href:
                logo_frame.addElement(Image(href=logo_href))
                logo_p.addElement(logo_frame)
                logo_cell.addElement(logo_p)
            else:
                logo_cell.addElement(P(text=f"Logo error: {logo_path}"))
                messagebox.showwarning("Logo Error", f"Could not embed logo from {logo_path}. File might be invalid or unsupported.")
        except FileNotFoundError:
            logo_cell.addElement(P(text=f"Logo file not found: {logo_path}"))
            messagebox.showerror("Logo Error", f"Logo file not found at: {logo_path}")
        except Exception as e:
            logo_cell.addElement(P(text=f"Error adding logo: {e}"))
            messagebox.showerror("Logo Error", f"An error occurred while adding the logo: {e}")
    else:
        logo_cell.addElement(P(text="")) # Empty cell if no logo

    header_row.addElement(logo_cell)

    # --- Text Info Cell (Right Side) ---
    text_cell = TableCell()
    today = datetime.now().strftime("%d.%m.%Y")
    headline = P(text=f"Befundbericht zur Bewegungsanalyse vom {today}")
    text_cell.addElement(headline)
    
    patient_info1 = P(text=f"{patient_full_title} {patient_name}")
    patient_info2 = P(text=f"geboren am: {patient_dob}")
    text_cell.addElement(patient_info1)
    text_cell.addElement(patient_info2)
    
    header_row.addElement(text_cell)
    header_table.addElement(header_row)
    doc.text.addElement(header_table)

    # Add a blank paragraph after the header table for a bit of spacing before the main content
    doc.text.addElement(P(text=""))

    # --- Aspect Ratio Calculation for main image ---
    max_width_cm = 16.0
    if image_width_px > 0:
        aspect_ratio = image_height_px / image_width_px
    else:
        aspect_ratio = 1
    frame_width_cm = max_width_cm
    frame_height_cm = max_width_cm * aspect_ratio

    # --- More Robust Image Insertion ---
    # Create a paragraph to hold the image frame
    p = P()
    doc.text.addElement(p)

    # Create the frame for the image
    frame = Frame(width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm")
    
    try:
        href = doc.addPicture(image_path)
        if href:
            frame.addElement(Image(href=href))
            p.addElement(frame)
        else:
            # Add a placeholder text if image adding fails
            p.addText("Image could not be added.")
    except Exception as e:
        # Also add a placeholder on exception
        p.addText(f"Error adding image: {e}")

    doc.save(odt_path)


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

class GenderSelector:
    def __init__(self, parent):
        self.parent = parent
        self.gender = None
        self.top = tk.Toplevel(parent)
        self.top.title("Select Patient Gender")
        self.top.transient(parent) # Make it a transient window for the parent
        self.top.grab_set() # Grab all events until this window is destroyed

        tk.Label(self.top, text="Select Patient Gender:").pack(pady=10)

        self.gender_var = tk.StringVar(value="Male") # Default selection
        
        male_radio = tk.Radiobutton(self.top, text="Male", variable=self.gender_var, value="Male")
        male_radio.pack(anchor="w", padx=20)
        
        female_radio = tk.Radiobutton(self.top, text="Female", variable=self.gender_var, value="Female")
        female_radio.pack(anchor="w", padx=20)

        button_frame = tk.Frame(self.top)
        button_frame.pack(pady=10)

        ok_button = tk.Button(button_frame, text="OK", command=self._on_ok)
        ok_button.pack(side="left", padx=5)

        cancel_button = tk.Button(button_frame, text="Cancel", command=self._on_cancel)
        cancel_button.pack(side="left", padx=5)

        # Center the window
        self.top.update_idletasks()
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = self.top.winfo_width()
        window_height = self.top.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"+{x}+{y}")

        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel) # Handle window close button

    def _on_ok(self):
        self.gender = self.gender_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.gender = None # Indicate cancellation
        self.top.destroy()

    def get_gender(self):
        self.parent.wait_window(self.top)
        return self.gender

class TitleSelector:
    def __init__(self, parent):
        self.parent = parent
        self.academic_title = None
        self.top = tk.Toplevel(parent)
        self.top.title("Select Academic Title")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="Select Academic Title:").pack(pady=10)

        self.title_var = tk.StringVar(value="None") # Default selection
        
        prof_dr_radio = tk.Radiobutton(self.top, text="Prof Dr.", variable=self.title_var, value="Prof Dr.")
        prof_dr_radio.pack(anchor="w", padx=20)
        
        dr_radio = tk.Radiobutton(self.top, text="Dr.", variable=self.title_var, value="Dr.")
        dr_radio.pack(anchor="w", padx=20)

        none_radio = tk.Radiobutton(self.top, text="None", variable=self.title_var, value="None")
        none_radio.pack(anchor="w", padx=20)

        button_frame = tk.Frame(self.top)
        button_frame.pack(pady=10)

        ok_button = tk.Button(button_frame, text="OK", command=self._on_ok)
        ok_button.pack(side="left", padx=5)

        cancel_button = tk.Button(button_frame, text="Cancel", command=self._on_cancel)
        cancel_button.pack(side="left", padx=5)

        self.top.update_idletasks()
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = self.top.winfo_width()
        window_height = self.top.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"+{x}+{y}")

        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_ok(self):
        self.academic_title = self.title_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.academic_title = None
        self.top.destroy()

    def get_title(self):
        self.parent.wait_window(self.top)
        return self.academic_title

def generate_report():
    # 1. Ask for patient gender using the new selector
    gender_selector = GenderSelector(root)
    gender = gender_selector.get_gender()

    if gender is None: # User cancelled gender selection
        messagebox.showinfo("Cancelled", "Gender selection was cancelled.")
        return
    
    # Determine salutation based on gender
    salutation = ""
    if gender == "Male":
        salutation = "Herr"
    elif gender == "Female":
        salutation = "Frau"

    # 2. Ask for academic title using the new selector
    title_selector = TitleSelector(root)
    academic_title = title_selector.get_title()

    if academic_title is None: # User cancelled title selection
        messagebox.showinfo("Cancelled", "Academic title selection was cancelled.")
        return
    
    # Construct the full title string to be used in the report
    patient_full_title = f"{salutation}"
    if academic_title and academic_title != "None":
        patient_full_title += f" {academic_title}"

    patient_name = simpledialog.askstring("Input", "Patient Name (Nachname, Vorname):")
    if not patient_name: return

    patient_dob = simpledialog.askstring("Input", "Geboren am (DD.MM.YYYY):")
    if not patient_dob: return

    # Hardcoded logo path
    logo_path = "/home/pm/Schreibtisch/Berichttool/logo_orthopassion.png"
    # No longer asking for logo path, it's permanently set

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
        
        # Use an absolute path for the temporary image to ensure it can be found
        image_path = os.path.abspath("temp_pdf_snippet.png")
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

        create_report(patient_full_title, patient_name, patient_dob, image_path, snippet_width, snippet_height, odt_path, logo_path)
        
        # Re-enable deleting the temp image file
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

print("Report generator closed. Goodbye, Carlo Lade, Hau rein, test!")
