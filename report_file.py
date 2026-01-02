import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
from odf.opendocument import OpenDocumentText
from odf.draw import Frame, Image
from odf.text import P
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.style import Style, TableColumnProperties, ParagraphProperties, TextProperties
from datetime import datetime
from PIL import Image as PILImage


def create_report(patient_full_title, patient_name, patient_dob, report_creator, odt_path, logo_path=None, second_logo_path=None):
    doc = OpenDocumentText()

    # Define styles for table columns
    logo_col_width_cm = 6.0
    text_col_width_cm = 12.0

    logo_col_style = Style(name="LogoColumnStyle", family="table-column")
    logo_col_style.addElement(TableColumnProperties(columnwidth=f"{logo_col_width_cm}cm"))
    doc.styles.addElement(logo_col_style)

    text_col_style = Style(name="TextColumnStyle", family="table-column")
    text_col_style.addElement(TableColumnProperties(columnwidth=f"{text_col_width_cm}cm"))
    doc.styles.addElement(text_col_style)

    # Define style for centered paragraphs
    center_p_style = Style(name="CenterParagraph", family="paragraph")
    center_p_style.addElement(ParagraphProperties(textalign="center"))
    doc.styles.addElement(center_p_style)

    # Define style for headings with page break (bold, 14pt)
    heading_with_break_style = Style(name="HeadingWithBreakStyle", family="paragraph")
    heading_with_break_style.addElement(ParagraphProperties(breakbefore="page"))
    heading_with_break_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
    doc.styles.addElement(heading_with_break_style)

    # Define style for headings (bold, 14pt)
    heading_style = Style(name="HeadingStyle", family="paragraph")
    heading_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
    doc.styles.addElement(heading_style)

    # Define style for background text (12pt font, justified, 1.5 line spacing)
    background_text_style = Style(name="BackgroundTextStyle", family="paragraph")
    background_text_style.addElement(ParagraphProperties(textalign="justify", lineheight="150%"))
    background_text_style.addElement(TextProperties(fontsize="12pt"))
    doc.styles.addElement(background_text_style)

    # Define style for page break
    page_break_style = Style(name="PageBreakStyle", family="paragraph")
    page_break_style.addElement(ParagraphProperties(breakbefore="page"))
    doc.styles.addElement(page_break_style)

    # Create header table
    header_table = Table(name="HeaderLayoutTable")
    header_table.addElement(TableColumn(stylename="LogoColumnStyle"))
    header_table.addElement(TableColumn(stylename="TextColumnStyle"))

    header_row = TableRow()

    # Logo Cell
    logo_cell = TableCell()
    if logo_path and os.path.exists(logo_path):
        try:
            with PILImage.open(logo_path) as logo_img:
                original_logo_width_px, original_logo_height_px = logo_img.size

            max_logo_width_cm = 5.0

            if original_logo_width_px > 0:
                aspect_ratio = original_logo_height_px / original_logo_width_px
            else:
                aspect_ratio = 1

            logo_frame_width_cm = min(max_logo_width_cm, logo_col_width_cm - 0.5)
            logo_frame_height_cm = logo_frame_width_cm * aspect_ratio

            logo_p = P()
            logo_frame = Frame(
                name="LogoFrame",
                width=f"{logo_frame_width_cm}cm",
                height=f"{logo_frame_height_cm}cm",
                anchortype="paragraph"
            )
            logo_href = doc.addPicture(logo_path)
            if logo_href:
                logo_frame.addElement(Image(href=logo_href))
                logo_p.addElement(logo_frame)
                logo_cell.addElement(logo_p)
            else:
                logo_cell.addElement(P(text=f"Logo error: {logo_path}"))
                messagebox.showwarning("Logo Error", f"Could not embed logo from {logo_path}")
        except FileNotFoundError:
            logo_cell.addElement(P(text=f"Logo file not found: {logo_path}"))
            messagebox.showerror("Logo Error", f"Logo file not found at: {logo_path}")
        except Exception as e:
            logo_cell.addElement(P(text=f"Error adding logo: {e}"))
            messagebox.showerror("Logo Error", f"An error occurred while adding the logo: {e}")
    else:
        logo_cell.addElement(P(text=""))

    header_row.addElement(logo_cell)

    # Text Info Cell
    text_cell = TableCell()
    today = datetime.now().strftime("%d.%m.%Y")
    headline = P(text=f"Befundbericht zur Bewegungsanalyse vom {today}", stylename="HeadingStyle")
    text_cell.addElement(headline)

    patient_info1 = P(text=f"{patient_full_title} {patient_name}")
    patient_info2 = P(text=f"geboren am: {patient_dob}")
    text_cell.addElement(patient_info1)
    text_cell.addElement(patient_info2)

    creator_info = P(text=f"Bericht erstellt von: {report_creator}")
    text_cell.addElement(creator_info)

    header_row.addElement(text_cell)
    header_table.addElement(header_row)
    doc.text.addElement(header_table)

    doc.text.addElement(P(text=""))

    # Second Logo (Centered)
    if second_logo_path and os.path.exists(second_logo_path):
        try:
            with PILImage.open(second_logo_path) as logo_img:
                original_logo_width_px, original_logo_height_px = logo_img.size

            max_second_logo_width_cm = 16.0
            if original_logo_width_px > 0:
                aspect_ratio = original_logo_height_px / original_logo_width_px
            else:
                aspect_ratio = 1

            second_logo_frame_width_cm = min(max_second_logo_width_cm, 18.0)
            second_logo_frame_height_cm = second_logo_frame_width_cm * aspect_ratio

            centered_p_for_logo = P(stylename="CenterParagraph")
            second_logo_frame = Frame(
                name="SecondLogoFrame",
                width=f"{second_logo_frame_width_cm}cm",
                height=f"{second_logo_frame_height_cm}cm",
                anchortype="paragraph"
            )
            second_logo_href = doc.addPicture(second_logo_path)
            if second_logo_href:
                second_logo_frame.addElement(Image(href=second_logo_href))
                centered_p_for_logo.addElement(second_logo_frame)
                doc.text.addElement(centered_p_for_logo)
            else:
                doc.text.addElement(P(text=f"Error adding second logo: {second_logo_path}"))
                messagebox.showwarning("Second Logo Error", f"Could not embed second logo from {second_logo_path}")
        except FileNotFoundError:
            doc.text.addElement(P(text=f"Second logo file not found: {second_logo_path}"))
            messagebox.showerror("Second Logo Error", f"Second logo file not found at: {second_logo_path}")
        except Exception as e:
            doc.text.addElement(P(text=f"Error adding second logo: {e}"))
            messagebox.showerror("Second Logo Error", f"An error occurred while adding the second logo: {e}")

    # Second Page Content - Hintergrund heading with page break
    doc.text.addElement(P(text="Hintergrund", stylename="HeadingWithBreakStyle"))

    # Add spacing after heading
    doc.text.addElement(P(text=""))

    background_text = """Vielen Dank, dass Sie sich für die Bewegungsanalyse bei uns entschieden haben. Durch die ganzheitliche Bewegungsanalyse besitzen Sie nun gute Voraussetzungen, um Ihre Beschwerden zu lindern. Denn ein Großteil aller orthopädischen Beschwerden sind auf Schonhaltungen und Kompensationsbewegungen (funktioneller Natur) zurückzuführen. Funktionelle Beschwerdeauslöser sind mit einer zielgerichteten Therapie gut zu behandeln. Die vorliegenden Analyseergebnisse dienen Ihrem Arzt oder Therapeuten für eine schnelle und effektive Entscheidungsfindung hinsichtlich Ihres Therapieplans. Nachfolgend möchten wir Ihnen wichtige Informationen zu unseren Messsystemen geben.
<BREAK>
Unsere 4D-Wirbelsäulenvermessung rekonstruiert mit einem strahlungsfreien Verfahren Ihre Rückenoberfläche und den Beckenstand dreidimensional über eine definierte Zeitspanne (vierte Dimension). Dazu wird ein Lichtraster auf Ihre Rückenoberfläche projiziert, über dessen Verzerrungen die funktionelle Stellung Ihrer Wirbelsäule errechnet wird. Die Vorteile der 4D-Wirbelsäulenvermessung sind vielfältig: keine Strahlenbelastung, geringe Messdauer, Darstellung ab- und aufsteigender Einflüsse auf den Körper (u.a. Zähne und Kiefergelenke, Füße), sehr sensitives Verfahren, synchrone Darstellung von Oberkörper, Becken, Beinachse und Fußdruckmessung in der Bewegung. Von besonderer Wichtigkeit ist die Möglichkeit, den Therapieerfolg über den Zeitverlauf ohne schädliche Strahlenbelastung sichtbar zu machen.
<BREAK>
Der Befundbericht der 4D-Wirbelsäulenvermessung dient der schriftlichen Zusammenfassung Ihrer Analyseergebnisse. Viele Auffälligkeiten stehen im wechselwirkenden Zusammenhang und sind immer wieder auf eine gemeinsame Ursache zurückzuführen. So kann z.B. eine Fehlhaltung der Wirbelsäule über Verkettungsmechanismen Auswirkungen auf die Beinachsen oder eine veränderte Ansteuerung  der Beinmuskulatur bewirken, die dann von der Norm abweicht und von uns dargestellt wird. Eine Auffälligkeit ist dabei nicht zwangsläufig negativ besetzt, sondern stellt erfolgreiche Ausgleichsversuche Ihres Körpers dar. Um langfristig Überlastungsschäden vorzubeugen, suchen wir gemeinsam mit Ihnen die Ursache für Beschwerden und Kompensationsmuster.
<BREAK>
Ausdrücklich ist darauf hinzuweisen, dass die Bewegungsanalyse für eine schlüssige Interpretation die Anamnese und klinische Untersuchung durch den Arzt nicht ersetzt, sondern ergänzt, und Ihnen ein differenziertes Therapiekonzept ermöglicht. Im Gegensatz zur 4D – Wirbelsäulenvermessung können diagnostische ergänzende bildgebende Verfahren (MRT, CT oder Röntgen) strukturelle Auffälligkeiten im Körper sichtbar machen."""

    text_segments = background_text.split("<BREAK>")
    for segment in text_segments:
        doc.text.addElement(P(text=segment.strip(), stylename="BackgroundTextStyle"))

    doc.save(odt_path)


class GenderSelector:
    def __init__(self, parent):
        self.parent = parent
        self.gender = None
        self.top = tk.Toplevel(parent)
        self.top.title("Select Patient Gender")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="Select Patient Gender:").pack(pady=10)

        self.gender_var = tk.StringVar(value="Male")

        tk.Radiobutton(self.top, text="Male", variable=self.gender_var, value="Male").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="Female", variable=self.gender_var, value="Female").pack(anchor="w", padx=20)

        button_frame = tk.Frame(self.top)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="OK", command=self._on_ok).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = self.top.winfo_width()
        window_height = self.top.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"+{x}+{y}")

    def _on_ok(self):
        self.gender = self.gender_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.gender = None
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

        self.title_var = tk.StringVar(value="None")

        tk.Radiobutton(self.top, text="Prof Dr.", variable=self.title_var, value="Prof Dr.").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="Dr.", variable=self.title_var, value="Dr.").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="None", variable=self.title_var, value="None").pack(anchor="w", padx=20)

        button_frame = tk.Frame(self.top)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="OK", command=self._on_ok).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = self.top.winfo_width()
        window_height = self.top.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"+{x}+{y}")

    def _on_ok(self):
        self.academic_title = self.title_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.academic_title = None
        self.top.destroy()

    def get_title(self):
        self.parent.wait_window(self.top)
        return self.academic_title


class ReportCreatorSelector:
    def __init__(self, parent):
        self.parent = parent
        self.creator_name = None
        self.top = tk.Toplevel(parent)
        self.top.title("Select Report Creator")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="Select the report creator:").pack(pady=10)

        self.creator_var = tk.StringVar(value="Carlo Lade")

        names = ["Carlo Lade", "Linnea Nirenberg", "Clara Guntermann", "Valentin Stark", "Lena Ranzinger"]
        for name in names:
            tk.Radiobutton(self.top, text=name, variable=self.creator_var, value=name).pack(anchor="w", padx=20)

        button_frame = tk.Frame(self.top)
        button_frame.pack(pady=10)

        tk.Button(button_frame, text="OK", command=self._on_ok).pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel).pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = self.top.winfo_width()
        window_height = self.top.winfo_height()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"+{x}+{y}")

    def _on_ok(self):
        self.creator_name = self.creator_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.creator_name = None
        self.top.destroy()

    def get_creator(self):
        self.parent.wait_window(self.top)
        return self.creator_name


def generate_report():
    gender_selector = GenderSelector(root)
    gender = gender_selector.get_gender()

    if gender is None:
        messagebox.showinfo("Cancelled", "Gender selection was cancelled.")
        return

    salutation = "Herr" if gender == "Male" else "Frau"

    title_selector = TitleSelector(root)
    academic_title = title_selector.get_title()

    if academic_title is None:
        messagebox.showinfo("Cancelled", "Academic title selection was cancelled.")
        return

    patient_full_title = salutation
    if academic_title and academic_title != "None":
        patient_full_title += f" {academic_title}"

    patient_name = simpledialog.askstring("Input", "Patient Name (Nachname, Vorname):")
    if not patient_name:
        return

    patient_dob = simpledialog.askstring("Input", "Geboren am (DD.MM.YYYY):")
    if not patient_dob:
        return

    creator_selector = ReportCreatorSelector(root)
    report_creator = creator_selector.get_creator()
    if report_creator is None:
        messagebox.showinfo("Cancelled", "Report creator selection was cancelled.")
        return

    logo_path = "/home/pm/Schreibtisch/Berichttool/logo_orthopassion.png"
    second_logo_path = "/home/pm/Schreibtisch/Berichttool/logo_2_orthopassion.png"

    odt_path = filedialog.asksaveasfilename(
        title="Save ODT report as",
        defaultextension=".odt",
        filetypes=[("ODT files", "*.odt")]
    )
    if not odt_path:
        return

    try:
        create_report(patient_full_title, patient_name, patient_dob, report_creator, odt_path, logo_path, second_logo_path)
        messagebox.showinfo("Success", f"Successfully created {odt_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Main application
root = tk.Tk()
root.title("Report Generator")

tk.Label(root, text="Generate a new report.").pack(pady=10)
tk.Button(root, text="Generate Report", command=generate_report).pack(pady=20)

root.mainloop()