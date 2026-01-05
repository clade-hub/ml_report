import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
from odf.opendocument import OpenDocumentText
from odf.draw import Frame, Image
from odf.text import P, List, ListItem
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.style import Style, TableColumnProperties, ParagraphProperties, TextProperties, ListLevelProperties
from odf.text import ListStyle, ListLevelStyleBullet
from datetime import datetime
from PIL import Image as PILImage, ImageTk
from pdf2image import convert_from_path
import tempfile


def crop_pdf_screenshot(pdf_path):
    """
    Convert first page of PDF to high-quality image and crop using hardcoded coordinates.
    Returns path to temporary cropped image file.
    """
    try:
        # Hardcoded crop coordinates as percentages
        LEFT_PCT = 1.14
        TOP_PCT = 24.58
        RIGHT_PCT = 56.98
        BOTTOM_PCT = 85.74

        # Convert first page at 300 DPI for high quality
        print("Converting PDF to high-quality image (300 DPI)...")
        images = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
        if not images:
            return None

        image = images[0]
        img_width, img_height = image.size
        print(f"Full image size: {img_width}x{img_height} pixels at 300 DPI")

        # Calculate crop coordinates from percentages
        left = int((LEFT_PCT / 100) * img_width)
        top = int((TOP_PCT / 100) * img_height)
        right = int((RIGHT_PCT / 100) * img_width)
        bottom = int((BOTTOM_PCT / 100) * img_height)

        print(f"Cropping to: ({left}, {top}, {right}, {bottom})")

        # Crop the image
        cropped_image = image.crop((left, top, right, bottom))
        crop_width, crop_height = cropped_image.size
        print(f"Cropped image size: {crop_width}x{crop_height} pixels")

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        cropped_image.save(temp_file.name, 'PNG', quality=95)
        print(f"Saved cropped image to: {temp_file.name}")

        return temp_file.name

    except Exception as e:
        print(f"Error cropping PDF screenshot: {e}")
        import traceback
        traceback.print_exc()
        return None


def crop_statik_screenshots(pdf_path):
    """
    Crop all 7 screenshots from statik.pdf using hardcoded coordinates.
    Returns list of paths to temporary cropped image files.
    """
    try:
        # Hardcoded crop coordinates for all 7 screenshots
        screenshots_config = [
            {"page": 2, "left": 12.59, "top": 29.09, "right": 68.49, "bottom": 80.10},  # Screenshot 1
            {"page": 5, "left": 11.97, "top": 28.12, "right": 70.31, "bottom": 84.85},  # Screenshot 2
            {"page": 3, "left": 12.76, "top": 31.59, "right": 67.12, "bottom": 81.95},  # Screenshot 3
            {"page": 4, "left": 13.50, "top": 32.72, "right": 66.55, "bottom": 83.00},  # Screenshot 4
            {"page": 6, "left": 1.03, "top": 34.57, "right": 59.66, "bottom": 90.09},   # Screenshot 5
            {"page": 6, "left": 1.65, "top": 15.07, "right": 92.25, "bottom": 34.57},   # Screenshot 6
            {"page": 6, "left": 59.66, "top": 34.81, "right": 91.62, "bottom": 89.20},  # Screenshot 7
        ]

        screenshot_paths = []

        for i, config in enumerate(screenshots_config, 1):
            print(f"\nProcessing Screenshot {i} from page {config['page']}...")

            # Convert specific page at 300 DPI for high quality
            images = convert_from_path(pdf_path, dpi=300, first_page=config['page'], last_page=config['page'])
            if not images:
                print(f"Warning: Could not convert page {config['page']}")
                screenshot_paths.append(None)
                continue

            image = images[0]
            img_width, img_height = image.size

            # Calculate crop coordinates from percentages
            left = int((config['left'] / 100) * img_width)
            top = int((config['top'] / 100) * img_height)
            right = int((config['right'] / 100) * img_width)
            bottom = int((config['bottom'] / 100) * img_height)

            # Crop the image
            cropped_image = image.crop((left, top, right, bottom))
            crop_width, crop_height = cropped_image.size
            print(f"Screenshot {i} cropped: {crop_width}x{crop_height} pixels")

            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            cropped_image.save(temp_file.name, 'PNG', quality=95)
            screenshot_paths.append(temp_file.name)

        return screenshot_paths

    except Exception as e:
        print(f"Error cropping statik screenshots: {e}")
        import traceback
        traceback.print_exc()
        return []


def crop_kraft_screenshots(pdf_path):
    """
    Crop 4 screenshots from kraft.pdf using hardcoded coordinates.
    Returns list of paths to temporary cropped image files.
    """
    try:
        # Hardcoded crop coordinates for 4 screenshots
        screenshots_config = [
            {"page": 1, "left": 0.85, "top": 19.98, "right": 48.43, "bottom": 78.97},   # Screenshot 1 (Kraftanalyse rechts-links - centered)
            {"page": 1, "left": 8.09, "top": 81.47, "right": 82.68, "bottom": 89.36},   # Screenshot 2 (Kraftanalyse rechts-links - bottom)
            {"page": 2, "left": 2.28, "top": 18.45, "right": 48.95, "bottom": 76.87},   # Screenshot 3 (Kraftanalyse Antagonist-Agonist - centered)
            {"page": 2, "left": 11.00, "top": 80.74, "right": 78.01, "bottom": 89.20},  # Screenshot 4 (Kraftanalyse Antagonist-Agonist - bottom)
        ]

        screenshot_paths = []

        for i, config in enumerate(screenshots_config, 1):
            print(f"\nProcessing Kraft Screenshot {i} from page {config['page']}...")

            # Convert specific page at 300 DPI for high quality
            images = convert_from_path(pdf_path, dpi=300, first_page=config['page'], last_page=config['page'])
            if not images:
                print(f"Warning: Could not convert page {config['page']}")
                screenshot_paths.append(None)
                continue

            image = images[0]
            img_width, img_height = image.size

            # Calculate crop coordinates from percentages
            left = int((config['left'] / 100) * img_width)
            top = int((config['top'] / 100) * img_height)
            right = int((config['right'] / 100) * img_width)
            bottom = int((config['bottom'] / 100) * img_height)

            # Crop the image
            cropped_image = image.crop((left, top, right, bottom))
            crop_width, crop_height = cropped_image.size
            print(f"Kraft Screenshot {i} cropped: {crop_width}x{crop_height} pixels")

            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            cropped_image.save(temp_file.name, 'PNG', quality=95)
            screenshot_paths.append(temp_file.name)

        return screenshot_paths

    except Exception as e:
        print(f"Error cropping kraft screenshots: {e}")
        import traceback
        traceback.print_exc()
        return []


def crop_gehen_screenshots(pdf_path):
    """
    Crop 8 screenshots from gehen.pdf using hardcoded coordinates.
    Returns list of paths to temporary cropped image files.
    """
    try:
        # Hardcoded crop coordinates for 8 screenshots
        screenshots_config = [
            {"page": 1, "left": 33.05, "top": 16.20, "right": 90.48, "bottom": 86.87},  # Screenshot 1 (Dynamische Beckenanalyse)
            {"page": 2, "left": 0.68, "top": 14.99, "right": 46.55, "bottom": 87.00},   # Screenshot 2 (Dynamische Wirbelsäulenanalyse)
            {"page": 3, "left": 16.87, "top": 18.05, "right": 91.17, "bottom": 73.25},  # Screenshot 3 (Ganganalyse - page 3)
            {"page": 5, "left": 13.73, "top": 18.69, "right": 92.08, "bottom": 80.82},  # Screenshot 4 (Ganganalyse - page 5)
            {"page": 4, "left": 2.62, "top": 17.57, "right": 92.88, "bottom": 87.11},   # Screenshot 5 (Ganganalyse - page 4)
            {"page": 6, "left": 2.74, "top": 22.24, "right": 93.56, "bottom": 83.80},   # Screenshot 6 (Ganganalyse - page 6)
            {"page": 8, "left": 3.13, "top": 19.10, "right": 94.02, "bottom": 87.11},   # Screenshot 7 (Dynamische Pedografie - page 8)
            {"page": 7, "left": 14.36, "top": 18.69, "right": 93.90, "bottom": 84.05},  # Screenshot 8 (Dynamische Pedografie - page 7)
        ]

        screenshot_paths = []

        for i, config in enumerate(screenshots_config, 1):
            print(f"\nProcessing Gehen Screenshot {i} from page {config['page']}...")

            # Convert specific page at 300 DPI for high quality
            images = convert_from_path(pdf_path, dpi=300, first_page=config['page'], last_page=config['page'])
            if not images:
                print(f"Warning: Could not convert page {config['page']}")
                screenshot_paths.append(None)
                continue

            image = images[0]
            img_width, img_height = image.size

            # Calculate crop coordinates from percentages
            left = int((config['left'] / 100) * img_width)
            top = int((config['top'] / 100) * img_height)
            right = int((config['right'] / 100) * img_width)
            bottom = int((config['bottom'] / 100) * img_height)

            # Crop the image
            cropped_image = image.crop((left, top, right, bottom))
            crop_width, crop_height = cropped_image.size
            print(f"Gehen Screenshot {i} cropped: {crop_width}x{crop_height} pixels")

            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            cropped_image.save(temp_file.name, 'PNG', quality=95)
            screenshot_paths.append(temp_file.name)

        return screenshot_paths

    except Exception as e:
        print(f"Error cropping gehen screenshots: {e}")
        import traceback
        traceback.print_exc()
        return []


def parse_ini_file(ini_path):
    """Parse the .ini file and extract kyphosis, lordosis, and scoliosis data"""
    try:
        with open(ini_path, 'r', encoding='utf-16') as f:
            lines = f.readlines()

        kyphosis_angle = None
        lordosis_angle = None
        scoliosis_angle = None
        surface_rotation_left = None
        surface_rotation_right = None
        lateral_deviation_left = None
        lateral_deviation_right = None
        sva_axis = None
        beckenhochstand = None

        # Line 8 for SVA axis (index 7)
        if len(lines) > 7:
            parts = lines[7].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    sva_axis = float(value_str)
                except ValueError:
                    pass

        # Line 11 for Beckenhochstand (index 10)
        if len(lines) > 10:
            parts = lines[10].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    beckenhochstand = float(value_str)
                except ValueError:
                    pass

        # Line 25 for kyphosis (index 24)
        if len(lines) > 24:
            parts = lines[24].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    kyphosis_angle = float(value_str)
                except ValueError:
                    pass

        # Line 28 for lordosis (index 27)
        if len(lines) > 27:
            parts = lines[27].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    lordosis_angle = float(value_str)
                except ValueError:
                    pass

        # Line 119 for scoliosis (index 118)
        if len(lines) > 118:
            parts = lines[118].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    scoliosis_angle = float(value_str)
                except ValueError:
                    pass

        # Line 34 for surface rotation left (index 33)
        if len(lines) > 33:
            parts = lines[33].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    surface_rotation_left = float(value_str)
                except ValueError:
                    pass

        # Line 35 for surface rotation right (index 34)
        if len(lines) > 34:
            parts = lines[34].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    surface_rotation_right = float(value_str)
                except ValueError:
                    pass

        # Line 40 for lateral deviation right (index 39)
        if len(lines) > 39:
            parts = lines[39].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    lateral_deviation_right = float(value_str)
                except ValueError:
                    pass

        # Line 41 for lateral deviation left (index 40)
        if len(lines) > 40:
            parts = lines[40].split('\t')
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    lateral_deviation_left = float(value_str)
                except ValueError:
                    pass

        return (kyphosis_angle, lordosis_angle, scoliosis_angle,
                surface_rotation_left, surface_rotation_right,
                lateral_deviation_left, lateral_deviation_right, sva_axis, beckenhochstand)
    except Exception as e:
        messagebox.showerror("Error", f"Error parsing INI file: {e}")
        return None, None, None, None, None, None, None, None, None


def parse_motion_ini_file(ini_path):
    """Parse the 4D motion .ini file and extract pelvic drop min and max values"""
    try:
        with open(ini_path, 'r', encoding='utf-16') as f:
            lines = f.readlines()

        pelvic_drop_min = None
        pelvic_drop_max = None

        # Line 11 (index 10) for pelvic drop min (column 3) and max (column 4)
        if len(lines) > 10:
            parts = lines[10].split('\t')
            # Column 3 (index 2) for min
            if len(parts) >= 3:
                value_str = parts[2].strip().replace(',', '.')
                try:
                    pelvic_drop_min = float(value_str)
                except ValueError:
                    pass
            # Column 4 (index 3) for max
            if len(parts) >= 4:
                value_str = parts[3].strip().replace(',', '.')
                try:
                    pelvic_drop_max = float(value_str)
                except ValueError:
                    pass

        return pelvic_drop_min, pelvic_drop_max
    except Exception as e:
        print(f"Error parsing motion INI file: {e}")
        return None, None


def calculate_pelvic_drop_sentence(beckenhochstand, beckenhochstand_side, motion_min, motion_max):
    """Calculate pelvic drop values and generate appropriate sentence

    Args:
        beckenhochstand: Value from 4D average .ini file
        beckenhochstand_side: "rechts", "links", or "gerade"
        motion_min: Min value from 4D motion .ini file
        motion_max: Max value from 4D motion .ini file

    Returns:
        Sentence describing the pelvic drop
    """
    if motion_min is None or motion_max is None:
        return None

    # Perform calculation based on Beckenhochstand side
    if beckenhochstand_side == "links" and beckenhochstand is not None:
        min_final = motion_min + beckenhochstand
        max_final = motion_max + beckenhochstand
    elif beckenhochstand_side == "rechts" and beckenhochstand is not None:
        min_final = motion_min - beckenhochstand
        max_final = motion_max - beckenhochstand
    else:  # "gerade"
        min_final = motion_min
        max_final = motion_max

    # Calculate absolute difference
    abs_difference = abs(abs(max_final) - abs(min_final))

    # Format values with one decimal place (using absolute values to remove minus sign)
    min_str = f"{abs(min_final):.1f}"
    max_str = f"{abs(max_final):.1f}"

    # Generate sentence based on difference
    if abs_difference < 2:
        # Symmetric pelvic drop
        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet ein symmetrischer Pelvic Drop von ({min_str}°R | {max_str}°L) statt"
    else:
        # Asymmetric pelvic drop - determine which stance phase
        if abs(max_final) > abs(min_final):
            stance_phase = "rechten"
        else:
            stance_phase = "linken"

        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet während der {stance_phase} Standphase ein asymmetrischer Pelvic Drop von ({min_str}°R | {max_str}°L) statt"

    return sentence


def classify_kyphosis(angle):
    """Classify kyphosis angle (same for both genders)"""
    if angle < 31:
        return "Eine Hypokyphose"
    elif 31 <= angle <= 38:
        return "Eine Tendenz zur Hypokyphose"
    elif 39 <= angle <= 56:
        return "Ein normgerechter Kyphosewinkel"
    elif 57 <= angle <= 64:
        return "Eine Tendenz zur Hyperkyphose"
    else:  # >= 64
        return "Eine Hyperkyphose"


def classify_lordosis(angle, gender):
    """Classify lordosis angle based on gender"""
    if gender == "Female":
        if angle < 24:
            return "Eine Hypolordose"
        elif 24 <= angle <= 32:
            return "Eine Tendenz zur Hypolordose"
        elif 33 <= angle <= 51:
            return "Ein normgerechter Lordosewinkel"
        elif 52 <= angle <= 60:
            return "Eine Tendenz zur Hyperlordose"
        else:  # > 60
            return "Eine Hyperlordose"
    else:  # Male
        if angle < 18:
            return "Eine Hypolordose"
        elif 18 <= angle <= 25:
            return "Eine Tendenz zur Hypolordose"
        elif 26 <= angle <= 42:
            return "Ein normgerechter Lordosewinkel"
        elif 43 <= angle <= 49:
            return "Eine Tendenz zur Hyperlordose"
        else:  # > 49
            return "Eine Hyperlordose"


def classify_scoliosis(angle, surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right):
    """Classify scoliosis angle and add surface rotation and lateral deviation"""
    # Base classification
    if angle < 10:
        base_text = "Eine skoliotische Haltungsstörung"
    else:
        base_text = "Eine Skoliose"

    # Determine surface rotation direction and value
    abs_left_rot = abs(surf_rot_left) if surf_rot_left is not None else 0
    abs_right_rot = abs(surf_rot_right) if surf_rot_right is not None else 0

    if abs_left_rot > abs_right_rot:
        rotation_direction = "links"
        rotation_value = abs_left_rot
    else:
        rotation_direction = "rechts"
        rotation_value = abs_right_rot

    # Determine lateral deviation direction and value
    abs_left_dev = abs(lat_dev_left) if lat_dev_left is not None else 0
    abs_right_dev = abs(lat_dev_right) if lat_dev_right is not None else 0

    if abs_left_dev > abs_right_dev:
        deviation_direction = "links"
        deviation_value = abs_left_dev
    else:
        deviation_direction = "rechts"
        deviation_value = abs_right_dev

    # Build complete sentence
    full_text = f"{base_text} mit vermehrter Oberflächenrotation ({rotation_value:.1f}°) nach {rotation_direction} und Seitabweichung ({deviation_value:.0f}mm) nach {deviation_direction}"

    return full_text


def generate_sim_sentence(isg_right, isg_left):
    """Generate SIM measurement sentence based on ISG blockage"""
    if isg_right == "Frei" and isg_left == "Frei":
        return "Der Beckenhochstand lässt sich im dynamischen Provokationstest auf der Simulationsplattform mit 10mm Ausgleich re/li beeinflussen, wodurch keine Fixierung im LWS/ISG Bereich festzustellen ist"

    # Determine which sides are blocked
    blocked_sides = []
    if isg_right == "blockiert":
        blocked_sides.append("rechts")
    if isg_left == "blockiert":
        blocked_sides.append("links")

    if len(blocked_sides) == 2:
        sides_text = "rechts und links"
    else:
        sides_text = blocked_sides[0]

    return f"Der Beckenhochstand lässt sich im dynamischen Provokationstest auf der Simulationsplattform mit 10mm Ausgleich {sides_text} nicht beeinflussen, was beweisend für eine Fixierung im LWS/ISG Bereich ist"


def generate_marker_sentence(markers):
    """Generate marker placement sentence based on selection"""
    if markers is None or markers.get('keine', False):
        return None

    selected_markers = []
    if markers.get('dl_dr', False):
        selected_markers.append("DL & DR")
    if markers.get('ws', False):
        selected_markers.append("WS")
    if markers.get('vp', False):
        selected_markers.append("VP")

    if not selected_markers:
        return None

    markers_text = ", ".join(selected_markers)
    return f"Zusätzliche Marker wurden geklebt: {markers_text}"


def create_report(patient_full_title, patient_name, patient_dob, report_creator, odt_path, gender, ini_path,
                  sim_performed, isg_right, isg_left, markers, logo_path=None, second_logo_path=None,
                  screenshot_path=None, statik_screenshots=None, pelvic_drop_sentence=None, gehen_screenshots=None, kraft_screenshots=None):
    doc = OpenDocumentText()

    # Define styles for table columns (3 columns: logo, text1, text2)
    logo_col_width_cm = 4.5
    text_col1_width_cm = 6.75
    text_col2_width_cm = 6.75

    logo_col_style = Style(name="LogoColumnStyle", family="table-column")
    logo_col_style.addElement(TableColumnProperties(columnwidth=f"{logo_col_width_cm}cm"))
    doc.styles.addElement(logo_col_style)

    text_col1_style = Style(name="TextColumn1Style", family="table-column")
    text_col1_style.addElement(TableColumnProperties(columnwidth=f"{text_col1_width_cm}cm"))
    doc.styles.addElement(text_col1_style)

    text_col2_style = Style(name="TextColumn2Style", family="table-column")
    text_col2_style.addElement(TableColumnProperties(columnwidth=f"{text_col2_width_cm}cm"))
    doc.styles.addElement(text_col2_style)

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

    # Define style for bullet list items (12pt font, 1.5 line spacing)
    bullet_text_style = Style(name="BulletTextStyle", family="paragraph")
    bullet_text_style.addElement(ParagraphProperties(lineheight="150%"))
    bullet_text_style.addElement(TextProperties(fontsize="12pt"))
    doc.styles.addElement(bullet_text_style)

    # Define bullet list style
    bullet_list_style = ListStyle(name="BulletList")
    bullet_level = ListLevelStyleBullet(level=1, stylename="BulletList", bulletchar="•")
    bullet_list_style.addElement(bullet_level)
    doc.styles.addElement(bullet_list_style)

    # Define style for page break
    page_break_style = Style(name="PageBreakStyle", family="paragraph")
    page_break_style.addElement(ParagraphProperties(breakbefore="page"))
    doc.styles.addElement(page_break_style)

    # Create header table with 3 columns
    header_table = Table(name="HeaderLayoutTable")
    header_table.addElement(TableColumn(stylename="LogoColumnStyle"))
    header_table.addElement(TableColumn(stylename="TextColumn1Style"))
    header_table.addElement(TableColumn(stylename="TextColumn2Style"))

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

    # Text Info Cell (spans 2 columns)
    text_cell = TableCell(numbercolumnsspanned=2)
    today = datetime.now().strftime("%d.%m.%Y")
    headline = P(text=f"Befundbericht zur Bewegungsanalyse vom {today}", stylename="HeadingStyle")
    text_cell.addElement(headline)

    patient_info1 = P(text=f"{patient_full_title} {patient_name}")
    patient_info2 = P(text=f"geboren am: {patient_dob}")
    text_cell.addElement(patient_info1)
    text_cell.addElement(patient_info2)

    creator_info = P(text=f"Bericht erstellt von: {report_creator} am {today}")
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

    # Third Page Content - Static Analysis
    doc.text.addElement(P(text="Statische Analyse", stylename="HeadingWithBreakStyle"))
    doc.text.addElement(P(text=""))

    # Parse INI file and generate analysis
    (kyphosis_angle, lordosis_angle, scoliosis_angle,
     surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right, sva_axis, beckenhochstand) = parse_ini_file(ini_path)

    if kyphosis_angle is not None and lordosis_angle is not None and scoliosis_angle is not None:
        # Create bullet list
        bullet_list = List(stylename="BulletList")

        # Kyphosis analysis
        kyphosis_text = classify_kyphosis(kyphosis_angle)
        kyphosis_item = ListItem()
        kyphosis_item.addElement(P(text=kyphosis_text, stylename="BulletTextStyle"))
        bullet_list.addElement(kyphosis_item)

        # Lordosis analysis
        lordosis_text = classify_lordosis(lordosis_angle, gender)
        lordosis_item = ListItem()
        lordosis_item.addElement(P(text=lordosis_text, stylename="BulletTextStyle"))
        bullet_list.addElement(lordosis_item)

        # Scoliosis analysis
        scoliosis_text = classify_scoliosis(scoliosis_angle, surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right)
        scoliosis_item = ListItem()
        scoliosis_item.addElement(P(text=scoliosis_text, stylename="BulletTextStyle"))
        bullet_list.addElement(scoliosis_item)

        # SVA axis analysis (only if above 50mm)
        if sva_axis is not None and sva_axis > 50:
            sva_text = f"Sagittale Dysbalance durch vermehrte anteriore Rumpfneigung ({sva_axis:.0f}mm); Norm <50mm"
            sva_item = ListItem()
            sva_item.addElement(P(text=sva_text, stylename="BulletTextStyle"))
            bullet_list.addElement(sva_item)

        # SIM measurement analysis (only if performed)
        if sim_performed == "Ja" and isg_right is not None and isg_left is not None:
            sim_text = generate_sim_sentence(isg_right, isg_left)
            sim_item = ListItem()
            sim_item.addElement(P(text=sim_text, stylename="BulletTextStyle"))
            bullet_list.addElement(sim_item)

        # Marker placement analysis (only if markers were placed)
        marker_text = generate_marker_sentence(markers)
        if marker_text is not None:
            marker_item = ListItem()
            marker_item.addElement(P(text=marker_text, stylename="BulletTextStyle"))
            bullet_list.addElement(marker_item)

        doc.text.addElement(bullet_list)

        # Add screenshot below bullet list if available
        if screenshot_path and os.path.exists(screenshot_path):
            try:
                # Add spacing before screenshot
                doc.text.addElement(P(text=""))

                # Get screenshot dimensions
                with PILImage.open(screenshot_path) as screenshot_img:
                    screenshot_width_px, screenshot_height_px = screenshot_img.size

                # Set image width to 16cm and calculate height maintaining aspect ratio
                screenshot_frame_width_cm = 16.0
                if screenshot_width_px > 0:
                    aspect_ratio = screenshot_height_px / screenshot_width_px
                else:
                    aspect_ratio = 1

                screenshot_frame_height_cm = screenshot_frame_width_cm * aspect_ratio

                # Create centered paragraph for screenshot
                centered_p_screenshot = P(stylename="CenterParagraph")
                screenshot_frame = Frame(
                    name="ScreenshotFrame",
                    width=f"{screenshot_frame_width_cm}cm",
                    height=f"{screenshot_frame_height_cm}cm",
                    anchortype="paragraph"
                )
                screenshot_href = doc.addPicture(screenshot_path)
                if screenshot_href:
                    screenshot_frame.addElement(Image(href=screenshot_href))
                    centered_p_screenshot.addElement(screenshot_frame)
                    doc.text.addElement(centered_p_screenshot)
                    print(f"Screenshot inserted: {screenshot_frame_width_cm}cm x {screenshot_frame_height_cm}cm")
                else:
                    print(f"Warning: Could not embed screenshot from {screenshot_path}")

            except Exception as e:
                print(f"Error adding screenshot: {e}")
                import traceback
                traceback.print_exc()

        # Add new section: Statische Beinachsen- und Haltungsanalyse
        if statik_screenshots and len(statik_screenshots) >= 7:
            # Page break before new section
            doc.text.addElement(P(text="Statische Beinachsen- und Haltungsanalyse", stylename="HeadingWithBreakStyle"))
            doc.text.addElement(P(text=""))

            # Add 3 bullet points with XXX placeholder
            statik_bullet_list = List(stylename="BulletList")

            for i in range(3):
                bullet_item = ListItem()
                bullet_item.addElement(P(text="XXX", stylename="BulletTextStyle"))
                statik_bullet_list.addElement(bullet_item)

            doc.text.addElement(statik_bullet_list)
            doc.text.addElement(P(text=""))

            # Add Screenshots 1 and 2 (16cm width, centered)
            for screenshot_idx in [0, 1]:  # Screenshots 1 and 2
                if statik_screenshots[screenshot_idx] and os.path.exists(statik_screenshots[screenshot_idx]):
                    try:
                        with PILImage.open(statik_screenshots[screenshot_idx]) as img:
                            img_width_px, img_height_px = img.size

                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio

                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(
                            name=f"StatikScreenshot{screenshot_idx + 1}",
                            width=f"{frame_width_cm}cm",
                            height=f"{frame_height_cm}cm",
                            anchortype="paragraph"
                        )
                        href = doc.addPicture(statik_screenshots[screenshot_idx])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            # No spacing - screenshots should be directly underneath each other
                            print(f"Statik Screenshot {screenshot_idx + 1} inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding statik screenshot {screenshot_idx + 1}: {e}")

            # Page break before Screenshots 3 and 4
            doc.text.addElement(P(text="", stylename="PageBreakStyle"))

            # Add Screenshots 3 and 4 (16cm width, centered)
            for screenshot_idx in [2, 3]:  # Screenshots 3 and 4
                if statik_screenshots[screenshot_idx] and os.path.exists(statik_screenshots[screenshot_idx]):
                    try:
                        with PILImage.open(statik_screenshots[screenshot_idx]) as img:
                            img_width_px, img_height_px = img.size

                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio

                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(
                            name=f"StatikScreenshot{screenshot_idx + 1}",
                            width=f"{frame_width_cm}cm",
                            height=f"{frame_height_cm}cm",
                            anchortype="paragraph"
                        )
                        href = doc.addPicture(statik_screenshots[screenshot_idx])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            # No spacing - screenshots should be directly underneath each other
                            print(f"Statik Screenshot {screenshot_idx + 1} inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding statik screenshot {screenshot_idx + 1}: {e}")

            # Page break before Statische Pedobarografie section
            doc.text.addElement(P(text="Statische Pedobarografie", stylename="HeadingWithBreakStyle"))

            # Add Screenshots 5, 6, 7 (16cm, 16cm, 8cm width, centered)
            screenshot_widths = [16.0, 16.0, 8.0]  # Widths for screenshots 5, 6, 7
            for i, screenshot_idx in enumerate([4, 5, 6]):  # Screenshots 5, 6, 7
                if statik_screenshots[screenshot_idx] and os.path.exists(statik_screenshots[screenshot_idx]):
                    try:
                        with PILImage.open(statik_screenshots[screenshot_idx]) as img:
                            img_width_px, img_height_px = img.size

                        frame_width_cm = screenshot_widths[i]
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio

                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(
                            name=f"StatikScreenshot{screenshot_idx + 1}",
                            width=f"{frame_width_cm}cm",
                            height=f"{frame_height_cm}cm",
                            anchortype="paragraph"
                        )
                        href = doc.addPicture(statik_screenshots[screenshot_idx])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            if i < 2:  # Add spacing after first two images
                                doc.text.addElement(P(text=""))
                            print(f"Statik Screenshot {screenshot_idx + 1} inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding statik screenshot {screenshot_idx + 1}: {e}")

        # Add new section: Dynamische Beckenanalyse
        if pelvic_drop_sentence:
            # Page break before new section
            doc.text.addElement(P(text="Dynamische Beckenanalyse", stylename="HeadingWithBreakStyle"))
            doc.text.addElement(P(text=""))

            # Add bullet point with pelvic drop sentence
            pelvic_bullet_list = List(stylename="BulletList")
            pelvic_item = ListItem()
            pelvic_item.addElement(P(text=pelvic_drop_sentence, stylename="BulletTextStyle"))
            pelvic_bullet_list.addElement(pelvic_item)
            doc.text.addElement(pelvic_bullet_list)

            # Add gehen screenshot 1 below pelvic drop sentence (if available)
            if gehen_screenshots and len(gehen_screenshots) >= 1 and gehen_screenshots[0] and os.path.exists(gehen_screenshots[0]):
                try:
                    # Add 2 empty lines as spacing
                    doc.text.addElement(P(text=""))
                    doc.text.addElement(P(text=""))

                    # Get screenshot dimensions
                    with PILImage.open(gehen_screenshots[0]) as img:
                        img_width_px, img_height_px = img.size

                    # Set width to 16cm and calculate height maintaining aspect ratio
                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    # Create centered paragraph for screenshot
                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="GehenScreenshot1",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(gehen_screenshots[0])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Gehen Screenshot 1 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")

                except Exception as e:
                    print(f"Error adding gehen screenshot 1: {e}")

        # Add new section: Dynamische Wirbelsäulenanalyse with gehen screenshot 2
        if gehen_screenshots and len(gehen_screenshots) >= 2 and gehen_screenshots[1] and os.path.exists(gehen_screenshots[1]):
            try:
                # Page break with heading
                doc.text.addElement(P(text="Dynamische Wirbelsäulenanalyse", stylename="HeadingWithBreakStyle"))

                # Get screenshot dimensions
                with PILImage.open(gehen_screenshots[1]) as img:
                    img_width_px, img_height_px = img.size

                # Set width to 16cm and calculate height maintaining aspect ratio
                frame_width_cm = 16.0
                aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                frame_height_cm = frame_width_cm * aspect_ratio

                # Create centered paragraph for screenshot
                centered_p = P(stylename="CenterParagraph")
                frame = Frame(
                    name="GehenScreenshot2",
                    width=f"{frame_width_cm}cm",
                    height=f"{frame_height_cm}cm",
                    anchortype="paragraph"
                )
                href = doc.addPicture(gehen_screenshots[1])
                if href:
                    frame.addElement(Image(href=href))
                    centered_p.addElement(frame)
                    doc.text.addElement(centered_p)
                    print(f"Gehen Screenshot 2 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")

            except Exception as e:
                print(f"Error adding gehen screenshot 2: {e}")

        # Add new section: Ganganalyse
        if gehen_screenshots and len(gehen_screenshots) >= 6:
            # Page break with heading
            doc.text.addElement(P(text="Ganganalyse", stylename="HeadingWithBreakStyle"))
            doc.text.addElement(P(text=""))

            # Add 4 bullet points with XXX placeholder
            ganganalyse_bullet_list = List(stylename="BulletList")
            for i in range(4):
                bullet_item = ListItem()
                bullet_item.addElement(P(text="XXX", stylename="BulletTextStyle"))
                ganganalyse_bullet_list.addElement(bullet_item)
            doc.text.addElement(ganganalyse_bullet_list)
            doc.text.addElement(P(text=""))

            # Add Screenshots 3 and 4 (indices 2 and 3) directly underneath each other
            for screenshot_idx in [2, 3]:  # Screenshots 3 and 4
                if gehen_screenshots[screenshot_idx] and os.path.exists(gehen_screenshots[screenshot_idx]):
                    try:
                        with PILImage.open(gehen_screenshots[screenshot_idx]) as img:
                            img_width_px, img_height_px = img.size

                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio

                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(
                            name=f"GehenScreenshot{screenshot_idx + 1}",
                            width=f"{frame_width_cm}cm",
                            height=f"{frame_height_cm}cm",
                            anchortype="paragraph"
                        )
                        href = doc.addPicture(gehen_screenshots[screenshot_idx])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Gehen Screenshot {screenshot_idx + 1} inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding gehen screenshot {screenshot_idx + 1}: {e}")

            # Page break before Screenshots 5 and 6
            doc.text.addElement(P(text="", stylename="PageBreakStyle"))

            # Add Screenshots 5 and 6 (indices 4 and 5) directly underneath each other
            for screenshot_idx in [4, 5]:  # Screenshots 5 and 6
                if gehen_screenshots[screenshot_idx] and os.path.exists(gehen_screenshots[screenshot_idx]):
                    try:
                        with PILImage.open(gehen_screenshots[screenshot_idx]) as img:
                            img_width_px, img_height_px = img.size

                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio

                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(
                            name=f"GehenScreenshot{screenshot_idx + 1}",
                            width=f"{frame_width_cm}cm",
                            height=f"{frame_height_cm}cm",
                            anchortype="paragraph"
                        )
                        href = doc.addPicture(gehen_screenshots[screenshot_idx])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Gehen Screenshot {screenshot_idx + 1} inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding gehen screenshot {screenshot_idx + 1}: {e}")

        # Add new section: Dynamische Pedografie
        if gehen_screenshots and len(gehen_screenshots) >= 8:
            # Page break with heading
            doc.text.addElement(P(text="Dynamische Pedografie", stylename="HeadingWithBreakStyle"))

            # Add Screenshot 7 (index 6)
            if gehen_screenshots[6] and os.path.exists(gehen_screenshots[6]):
                try:
                    with PILImage.open(gehen_screenshots[6]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="GehenScreenshot7",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(gehen_screenshots[6])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Gehen Screenshot 7 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding gehen screenshot 7: {e}")

            # Add 2 empty lines spacing
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))

            # Add Screenshot 8 (index 7)
            if gehen_screenshots[7] and os.path.exists(gehen_screenshots[7]):
                try:
                    with PILImage.open(gehen_screenshots[7]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="GehenScreenshot8",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(gehen_screenshots[7])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Gehen Screenshot 8 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding gehen screenshot 8: {e}")

        # Add new section: Kraftanalyse Vergleich rechts - links
        if kraft_screenshots and len(kraft_screenshots) >= 2:
            # Page break with heading
            doc.text.addElement(P(text="Kraftanalyse Vergleich rechts - links", stylename="HeadingWithBreakStyle"))

            # Add 1 empty line after heading
            doc.text.addElement(P(text=""))

            # Add Screenshot 1 (index 0) - centered horizontally and vertically
            if kraft_screenshots[0] and os.path.exists(kraft_screenshots[0]):
                try:
                    with PILImage.open(kraft_screenshots[0]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="KraftScreenshot1",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(kraft_screenshots[0])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Kraft Screenshot 1 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding kraft screenshot 1: {e}")

            # Add spacing to push screenshot 2 to the bottom
            for _ in range(10):
                doc.text.addElement(P(text=""))

            # Add Screenshot 2 (index 1) - at the bottom
            if kraft_screenshots[1] and os.path.exists(kraft_screenshots[1]):
                try:
                    with PILImage.open(kraft_screenshots[1]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="KraftScreenshot2",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(kraft_screenshots[1])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Kraft Screenshot 2 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding kraft screenshot 2: {e}")

        # Add new section: Kraftanalyse Vergleich Antagonist - Agonist
        if kraft_screenshots and len(kraft_screenshots) >= 4:
            # Page break with heading
            doc.text.addElement(P(text="Kraftanalyse Vergleich Antagonist - Agonist", stylename="HeadingWithBreakStyle"))

            # Add 1 empty line after heading
            doc.text.addElement(P(text=""))

            # Add Screenshot 3 (index 2) - centered horizontally and vertically
            if kraft_screenshots[2] and os.path.exists(kraft_screenshots[2]):
                try:
                    with PILImage.open(kraft_screenshots[2]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="KraftScreenshot3",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(kraft_screenshots[2])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Kraft Screenshot 3 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding kraft screenshot 3: {e}")

            # Add spacing to push screenshot 4 to the bottom
            for _ in range(10):
                doc.text.addElement(P(text=""))

            # Add Screenshot 4 (index 3) - at the bottom
            if kraft_screenshots[3] and os.path.exists(kraft_screenshots[3]):
                try:
                    with PILImage.open(kraft_screenshots[3]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="KraftScreenshot4",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(kraft_screenshots[3])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"Kraft Screenshot 4 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding kraft screenshot 4: {e}")

    else:
        doc.text.addElement(P(text="Fehler beim Lesen der Messwerte aus der INI-Datei."))

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


class SIMMeasurementSelector:
    def __init__(self, parent):
        self.parent = parent
        self.sim_performed = None
        self.top = tk.Toplevel(parent)
        self.top.title("SIM-Messung")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="SIM-Messung durchgeführt?").pack(pady=10)

        self.sim_var = tk.StringVar(value="Nein")

        tk.Radiobutton(self.top, text="Ja", variable=self.sim_var, value="Ja").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="Nein", variable=self.sim_var, value="Nein").pack(anchor="w", padx=20)

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
        self.sim_performed = self.sim_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.sim_performed = None
        self.top.destroy()

    def get_sim_status(self):
        self.parent.wait_window(self.top)
        return self.sim_performed


class ISGSelector:
    def __init__(self, parent):
        self.parent = parent
        self.isg_right = None
        self.isg_left = None
        self.top = tk.Toplevel(parent)
        self.top.title("ISG Status")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="ISG rechts:").pack(pady=5)
        self.right_var = tk.StringVar(value="Frei")
        tk.Radiobutton(self.top, text="Frei", variable=self.right_var, value="Frei").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="blockiert", variable=self.right_var, value="blockiert").pack(anchor="w", padx=20)

        tk.Label(self.top, text="ISG links:").pack(pady=5)
        self.left_var = tk.StringVar(value="Frei")
        tk.Radiobutton(self.top, text="Frei", variable=self.left_var, value="Frei").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="blockiert", variable=self.left_var, value="blockiert").pack(anchor="w", padx=20)

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
        self.isg_right = self.right_var.get()
        self.isg_left = self.left_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.isg_right = None
        self.isg_left = None
        self.top.destroy()

    def get_isg_status(self):
        self.parent.wait_window(self.top)
        return self.isg_right, self.isg_left


class MarkerSelector:
    def __init__(self, parent):
        self.parent = parent
        self.markers = None
        self.top = tk.Toplevel(parent)
        self.top.title("Marker Placement")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="Marker geklebt?").pack(pady=10)

        self.dl_dr_var = tk.BooleanVar(value=False)
        self.ws_var = tk.BooleanVar(value=False)
        self.vp_var = tk.BooleanVar(value=False)
        self.keine_var = tk.BooleanVar(value=False)

        tk.Checkbutton(self.top, text="DL & DR", variable=self.dl_dr_var).pack(anchor="w", padx=20)
        tk.Checkbutton(self.top, text="WS", variable=self.ws_var).pack(anchor="w", padx=20)
        tk.Checkbutton(self.top, text="VP", variable=self.vp_var).pack(anchor="w", padx=20)
        tk.Checkbutton(self.top, text="keine", variable=self.keine_var).pack(anchor="w", padx=20)

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
        self.markers = {
            'dl_dr': self.dl_dr_var.get(),
            'ws': self.ws_var.get(),
            'vp': self.vp_var.get(),
            'keine': self.keine_var.get()
        }
        self.top.destroy()

    def _on_cancel(self):
        self.markers = None
        self.top.destroy()

    def get_markers(self):
        self.parent.wait_window(self.top)
        return self.markers


class BeckenhochstandSelector:
    def __init__(self, parent):
        self.parent = parent
        self.beckenhochstand_side = None
        self.top = tk.Toplevel(parent)
        self.top.title("Beckenhochstand Position")
        self.top.transient(parent)
        self.top.grab_set()

        tk.Label(self.top, text="Für die Pelvic Drop Berechnung,\nist der Beckenhochstand rechts oder links oder gerade?",
                 font=("Arial", 10)).pack(pady=10)

        self.side_var = tk.StringVar(value="gerade")

        tk.Radiobutton(self.top, text="rechts", variable=self.side_var, value="rechts").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="links", variable=self.side_var, value="links").pack(anchor="w", padx=20)
        tk.Radiobutton(self.top, text="gerade", variable=self.side_var, value="gerade").pack(anchor="w", padx=20)

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
        self.beckenhochstand_side = self.side_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.beckenhochstand_side = None
        self.top.destroy()

    def get_beckenhochstand_side(self):
        self.parent.wait_window(self.top)
        return self.beckenhochstand_side


class CoordinateFinder:
    def __init__(self, parent):
        self.parent = parent
        self.image = None
        self.photo = None
        self.clicks = []
        self.canvas = None
        self.top = None

    def find_coordinates(self):
        # Ask user to select folder
        folder_path = filedialog.askdirectory(title="Select folder containing 4d_average.pdf")
        if not folder_path:
            messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
            return

        # Look for 4d_average.pdf
        pdf_path = os.path.join(folder_path, "4d_average.pdf")
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"Could not find 4d_average.pdf in {folder_path}")
            return

        self.find_coordinates_from_path(pdf_path)

    def find_coordinates_from_path(self, pdf_path, page_num=1, screenshot_name="Screenshot"):
        """Find coordinates from a specific page of a PDF

        Args:
            pdf_path: Path to the PDF file
            page_num: Page number to extract (default: 1)
            screenshot_name: Name to display in window title (default: "Screenshot")
        """
        try:
            # Convert specific page of PDF to image at 150 DPI (for display on laptop)
            print(f"Converting PDF page {page_num} to image...")
            images = convert_from_path(pdf_path, dpi=150, first_page=page_num, last_page=page_num)
            if not images:
                messagebox.showerror("Error", f"Could not convert PDF page {page_num} to image")
                return

            self.image = images[0]
            self.current_page = page_num
            self.screenshot_name = screenshot_name
            print(f"Image size: {self.image.size[0]}x{self.image.size[1]} pixels (at 150 DPI)")

            # Create window to display image
            self.top = tk.Toplevel(self.parent)
            window_title = f"{screenshot_name} - Page {page_num} - Click Top-Left, then Bottom-Right"
            self.top.title(window_title)
            self.top.transient(self.parent)
            self.top.grab_set()

            # Scale image if too large for screen
            max_width = 1200
            max_height = 800
            img_width, img_height = self.image.size

            scale = 1.0
            if img_width > max_width or img_height > max_height:
                scale_w = max_width / img_width
                scale_h = max_height / img_height
                scale = min(scale_w, scale_h)
                display_width = int(img_width * scale)
                display_height = int(img_height * scale)
                display_image = self.image.resize((display_width, display_height), PILImage.Resampling.LANCZOS)
            else:
                display_image = self.image
                display_width = img_width
                display_height = img_height

            self.scale = scale
            print(f"Display scale: {scale} (Display size: {display_width}x{display_height})")

            # Create canvas with image
            self.canvas = tk.Canvas(self.top, width=display_width, height=display_height)
            self.canvas.pack()

            # Convert PIL image to PhotoImage for display
            self.photo = ImageTk.PhotoImage(display_image)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

            # Add instructions
            instruction_text = "Click 1: Top-Left corner, Click 2: Bottom-Right corner"
            tk.Label(self.top, text=instruction_text, font=("Arial", 12, "bold")).pack(pady=5)

            # Bind click event
            self.canvas.bind("<Button-1>", self.on_click)

        except Exception as e:
            messagebox.showerror("Error", f"Error processing PDF: {e}")
            import traceback
            traceback.print_exc()

    def wait_for_completion(self):
        """Wait for the coordinate finder window to close"""
        if self.top:
            self.parent.wait_window(self.top)

    def on_click(self, event):
        # Get click coordinates (scaled to display)
        x_display = event.x
        y_display = event.y

        # Convert to original image coordinates
        x_original = int(x_display / self.scale)
        y_original = int(y_display / self.scale)

        # Store click
        self.clicks.append((x_original, y_original))

        # Draw marker on canvas
        r = 5  # radius
        self.canvas.create_oval(x_display-r, y_display-r, x_display+r, y_display+r,
                                fill="red", outline="white", width=2)
        self.canvas.create_text(x_display, y_display-15, text=f"Click {len(self.clicks)}",
                                fill="red", font=("Arial", 10, "bold"))

        print(f"Click {len(self.clicks)}: Display({x_display}, {y_display}) -> Original({x_original}, {y_original})")

        # If we have 2 clicks, calculate crop coordinates
        if len(self.clicks) == 2:
            self.calculate_coordinates()

    def calculate_coordinates(self):
        x1, y1 = self.clicks[0]
        x2, y2 = self.clicks[1]

        # Ensure top-left and bottom-right
        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)

        width = right - left
        height = bottom - top

        img_width, img_height = self.image.size

        # Calculate as percentages
        left_pct = (left / img_width) * 100
        top_pct = (top / img_height) * 100
        right_pct = (right / img_width) * 100
        bottom_pct = (bottom / img_height) * 100

        # Draw rectangle on canvas
        x1_display = int(left * self.scale)
        y1_display = int(top * self.scale)
        x2_display = int(right * self.scale)
        y2_display = int(bottom * self.scale)
        self.canvas.create_rectangle(x1_display, y1_display, x2_display, y2_display,
                                      outline="green", width=3)

        # Store final coordinates for programmatic access
        self.final_coordinates = {
            'left_pct': left_pct,
            'top_pct': top_pct,
            'right_pct': right_pct,
            'bottom_pct': bottom_pct,
            'left': left,
            'top': top,
            'right': right,
            'bottom': bottom,
            'width': width,
            'height': height
        }

        # Create result message
        result = f"""
CROP COORDINATES FOUND (at 150 DPI):
=====================================
Pixel coordinates (left, top, right, bottom):
({left}, {top}, {right}, {bottom})

Crop size: {width}x{height} pixels

Percentage of image (for any DPI):
Left: {left_pct:.2f}%
Top: {top_pct:.2f}%
Right: {right_pct:.2f}%
Bottom: {bottom_pct:.2f}%

For hardcoding, use these percentage values to calculate
coordinates at any DPI (e.g., 300 DPI for high quality).
"""

        print("\n" + "="*50)
        print(result)
        print("="*50 + "\n")

        messagebox.showinfo("Coordinates Found", result)

        self.top.destroy()


class StatikCoordinateFinder:
    """Coordinate finder for all 7 screenshots from statik.pdf"""
    def __init__(self, parent):
        self.parent = parent
        self.screenshots_config = [
            {"name": "Screenshot 1", "page": 2, "description": "Page 2 - 16cm width"},
            {"name": "Screenshot 2", "page": 5, "description": "Page 5 - 16cm width"},
            {"name": "Screenshot 3", "page": 3, "description": "Page 3 - 16cm width"},
            {"name": "Screenshot 4", "page": 4, "description": "Page 4 - 16cm width"},
            {"name": "Screenshot 5", "page": 6, "description": "Page 6 - 16cm width"},
            {"name": "Screenshot 6", "page": 6, "description": "Page 6 - 16cm width (different area)"},
            {"name": "Screenshot 7", "page": 6, "description": "Page 6 - 8cm width (half size)"},
        ]
        self.current_screenshot_index = 0
        self.all_coordinates = []

    def find_all_coordinates(self):
        """Start the coordinate finding process"""
        # Ask user to select folder
        folder_path = filedialog.askdirectory(title="Select folder containing statik.pdf")
        if not folder_path:
            messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
            return

        # Look for statik.pdf
        pdf_path = os.path.join(folder_path, "statik.pdf")
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"Could not find statik.pdf in {folder_path}")
            return

        self.pdf_path = pdf_path
        self.process_next_screenshot()

    def process_next_screenshot(self):
        """Process the next screenshot in the sequence"""
        if self.current_screenshot_index >= len(self.screenshots_config):
            # All screenshots processed
            self.show_all_results()
            return

        config = self.screenshots_config[self.current_screenshot_index]
        print(f"\n{'='*60}")
        print(f"Processing {config['name']}: {config['description']}")
        print(f"{'='*60}")

        # Create coordinate finder for this page
        finder = CoordinateFinder(self.parent)
        finder.find_coordinates_from_path(self.pdf_path, config['page'], config['name'])
        finder.wait_for_completion()

        # Store the coordinates if they were found
        if hasattr(finder, 'final_coordinates'):
            self.all_coordinates.append({
                'screenshot': config['name'],
                'page': config['page'],
                'coordinates': finder.final_coordinates
            })

        # Move to next screenshot
        self.current_screenshot_index += 1
        self.process_next_screenshot()

    def show_all_results(self):
        """Show all collected coordinates"""
        print("\n" + "="*70)
        print("ALL STATIK.PDF COORDINATES")
        print("="*70)

        result_text = "ALL STATIK.PDF COORDINATES:\n" + "="*70 + "\n\n"

        for coord_data in self.all_coordinates:
            coords = coord_data['coordinates']
            text = f"{coord_data['screenshot']} (Page {coord_data['page']}):\n"
            text += f"  Left: {coords['left_pct']:.2f}%\n"
            text += f"  Top: {coords['top_pct']:.2f}%\n"
            text += f"  Right: {coords['right_pct']:.2f}%\n"
            text += f"  Bottom: {coords['bottom_pct']:.2f}%\n\n"

            print(text)
            result_text += text

        messagebox.showinfo("All Coordinates Found", result_text)


class GehenCoordinateFinder:
    """Coordinate finder for new gehen.pdf screenshots (pages 3, 5, 4, 6, 8, 7)"""
    def __init__(self, parent):
        self.parent = parent
        self.screenshots_config = [
            {"name": "Ganganalyse Screenshot 1", "page": 3, "description": "Page 3 - 16cm width (Ganganalyse)"},
            {"name": "Ganganalyse Screenshot 2", "page": 5, "description": "Page 5 - 16cm width (Ganganalyse)"},
            {"name": "Ganganalyse Screenshot 3", "page": 4, "description": "Page 4 - 16cm width (Ganganalyse)"},
            {"name": "Ganganalyse Screenshot 4", "page": 6, "description": "Page 6 - 16cm width (Ganganalyse)"},
            {"name": "Dynamische Pedografie Screenshot 1", "page": 8, "description": "Page 8 - 16cm width (Dynamische Pedografie)"},
            {"name": "Dynamische Pedografie Screenshot 2", "page": 7, "description": "Page 7 - 16cm width (Dynamische Pedografie)"},
        ]
        self.current_screenshot_index = 0
        self.all_coordinates = []

    def find_all_coordinates(self):
        """Start the coordinate finding process"""
        # Ask user to select folder
        folder_path = filedialog.askdirectory(title="Select folder containing gehen.pdf")
        if not folder_path:
            messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
            return

        # Look for gehen.pdf
        pdf_path = os.path.join(folder_path, "gehen.pdf")
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"Could not find gehen.pdf in {folder_path}")
            return

        self.pdf_path = pdf_path
        self.process_next_screenshot()

    def process_next_screenshot(self):
        """Process the next screenshot in the sequence"""
        if self.current_screenshot_index >= len(self.screenshots_config):
            # All screenshots processed
            self.show_all_results()
            return

        config = self.screenshots_config[self.current_screenshot_index]
        print(f"\n{'='*60}")
        print(f"Processing {config['name']}: {config['description']}")
        print(f"{'='*60}")

        # Create coordinate finder for this page
        finder = CoordinateFinder(self.parent)
        finder.find_coordinates_from_path(self.pdf_path, config['page'], config['name'])
        finder.wait_for_completion()

        # Store the coordinates if they were found
        if hasattr(finder, 'final_coordinates'):
            self.all_coordinates.append({
                'screenshot': config['name'],
                'page': config['page'],
                'coordinates': finder.final_coordinates
            })

        # Move to next screenshot
        self.current_screenshot_index += 1
        self.process_next_screenshot()

    def show_all_results(self):
        """Show all collected coordinates"""
        print("\n" + "="*70)
        print("NEW GEHEN.PDF COORDINATES (Pages 3, 5, 4, 6, 8, 7)")
        print("="*70)

        result_text = "NEW GEHEN.PDF COORDINATES (Pages 3, 5, 4, 6, 8, 7):\n" + "="*70 + "\n\n"

        for coord_data in self.all_coordinates:
            coords = coord_data['coordinates']
            text = f"{coord_data['screenshot']} (Page {coord_data['page']}):\n"
            text += f"  Left: {coords['left_pct']:.2f}%\n"
            text += f"  Top: {coords['top_pct']:.2f}%\n"
            text += f"  Right: {coords['right_pct']:.2f}%\n"
            text += f"  Bottom: {coords['bottom_pct']:.2f}%\n\n"

            print(text)
            result_text += text

        messagebox.showinfo("All Coordinates Found", result_text)


class KraftCoordinateFinder:
    """Coordinate finder for kraft.pdf screenshots (4 screenshots from pages 1 and 2)"""
    def __init__(self, parent):
        self.parent = parent
        self.screenshots_config = [
            {"name": "Kraftanalyse rechts-links Screenshot 1", "page": 1, "description": "Page 1 - 16cm width (centered)"},
            {"name": "Kraftanalyse rechts-links Screenshot 2", "page": 1, "description": "Page 1 - 16cm width (bottom)"},
            {"name": "Kraftanalyse Antagonist-Agonist Screenshot 1", "page": 2, "description": "Page 2 - 16cm width (centered)"},
            {"name": "Kraftanalyse Antagonist-Agonist Screenshot 2", "page": 2, "description": "Page 2 - 16cm width (bottom)"},
        ]
        self.current_screenshot_index = 0
        self.all_coordinates = []

    def find_all_coordinates(self):
        """Start the coordinate finding process"""
        # Ask user to select folder
        folder_path = filedialog.askdirectory(title="Select folder containing kraft.pdf")
        if not folder_path:
            messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
            return

        # Look for kraft.pdf
        pdf_path = os.path.join(folder_path, "kraft.pdf")
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"Could not find kraft.pdf in {folder_path}")
            return

        self.pdf_path = pdf_path
        self.process_next_screenshot()

    def process_next_screenshot(self):
        """Process the next screenshot in the sequence"""
        if self.current_screenshot_index >= len(self.screenshots_config):
            # All screenshots processed
            self.show_all_results()
            return

        config = self.screenshots_config[self.current_screenshot_index]
        print(f"\n{'='*60}")
        print(f"Processing {config['name']}: {config['description']}")
        print(f"{'='*60}")

        # Create coordinate finder for this page
        finder = CoordinateFinder(self.parent)
        finder.find_coordinates_from_path(self.pdf_path, config['page'], config['name'])
        finder.wait_for_completion()

        # Store the coordinates if they were found
        if hasattr(finder, 'final_coordinates'):
            self.all_coordinates.append({
                'screenshot': config['name'],
                'page': config['page'],
                'coordinates': finder.final_coordinates
            })

        # Move to next screenshot
        self.current_screenshot_index += 1
        self.process_next_screenshot()

    def show_all_results(self):
        """Show all collected coordinates"""
        print("\n" + "="*70)
        print("KRAFT.PDF COORDINATES")
        print("="*70)

        result_text = "KRAFT.PDF COORDINATES:\n" + "="*70 + "\n\n"

        for coord_data in self.all_coordinates:
            coords = coord_data['coordinates']
            text = f"{coord_data['screenshot']} (Page {coord_data['page']}):\n"
            text += f"  Left: {coords['left_pct']:.2f}%\n"
            text += f"  Top: {coords['top_pct']:.2f}%\n"
            text += f"  Right: {coords['right_pct']:.2f}%\n"
            text += f"  Bottom: {coords['bottom_pct']:.2f}%\n\n"

            print(text)
            result_text += text

        messagebox.showinfo("All Coordinates Found", result_text)


def find_statik_coordinates():
    """Helper function to find all statik.pdf coordinates"""
    finder = StatikCoordinateFinder(root)
    finder.find_all_coordinates()


def find_gehen_coordinates():
    """Helper function to find all gehen.pdf coordinates"""
    finder = GehenCoordinateFinder(root)
    finder.find_all_coordinates()


def find_kraft_coordinates():
    """Helper function to find all kraft.pdf coordinates"""
    finder = KraftCoordinateFinder(root)
    finder.find_all_coordinates()


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

    # Ask for folder containing measurement files
    folder_path = filedialog.askdirectory(title="Select folder containing measurement files")
    if not folder_path:
        messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
        return

    # Auto-find .ini file containing "4D average" (case-insensitive)
    ini_path = None
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.ini') and '4d average' in filename.lower():
            ini_path = os.path.join(folder_path, filename)
            break

    if not ini_path:
        messagebox.showerror("Error", "Could not find .ini file containing '4D average' in the selected folder.")
        return

    print(f"Found .ini file: {ini_path}")

    # Auto-find 4d_average.pdf
    pdf_path = os.path.join(folder_path, "4d_average.pdf")
    if not os.path.exists(pdf_path):
        messagebox.showerror("Error", f"Could not find 4d_average.pdf in {folder_path}")
        return

    print(f"Found PDF file: {pdf_path}")

    # Crop screenshot from 4d_average.pdf
    screenshot_path = crop_pdf_screenshot(pdf_path)
    if not screenshot_path:
        messagebox.showerror("Error", "Failed to crop screenshot from PDF")
        return

    # Auto-find and crop screenshots from statik.pdf
    statik_pdf_path = os.path.join(folder_path, "statik.pdf")
    statik_screenshots = []
    if os.path.exists(statik_pdf_path):
        print(f"Found statik.pdf: {statik_pdf_path}")
        statik_screenshots = crop_statik_screenshots(statik_pdf_path)
        if not statik_screenshots or len(statik_screenshots) < 7:
            print("Warning: Failed to crop all statik screenshots")
    else:
        print("Warning: statik.pdf not found in folder")

    # Auto-find and crop screenshots from gehen.pdf
    gehen_pdf_path = os.path.join(folder_path, "gehen.pdf")
    gehen_screenshots = []
    if os.path.exists(gehen_pdf_path):
        print(f"Found gehen.pdf: {gehen_pdf_path}")
        gehen_screenshots = crop_gehen_screenshots(gehen_pdf_path)
        if not gehen_screenshots or len(gehen_screenshots) < 8:
            print("Warning: Failed to crop all gehen screenshots")
    else:
        print("Warning: gehen.pdf not found in folder")

    # Auto-find and crop screenshots from kraft.pdf
    kraft_pdf_path = os.path.join(folder_path, "kraft.pdf")
    kraft_screenshots = []
    if os.path.exists(kraft_pdf_path):
        print(f"Found kraft.pdf: {kraft_pdf_path}")
        kraft_screenshots = crop_kraft_screenshots(kraft_pdf_path)
        if not kraft_screenshots or len(kraft_screenshots) < 4:
            print("Warning: Failed to crop all kraft screenshots")
    else:
        print("Warning: kraft.pdf not found in folder")

    # Ask for SIM measurement
    sim_selector = SIMMeasurementSelector(root)
    sim_performed = sim_selector.get_sim_status()
    if sim_performed is None:
        messagebox.showinfo("Cancelled", "SIM measurement selection was cancelled.")
        return

    # If SIM was performed, ask for ISG status
    isg_right = None
    isg_left = None
    if sim_performed == "Ja":
        isg_selector = ISGSelector(root)
        isg_right, isg_left = isg_selector.get_isg_status()
        if isg_right is None or isg_left is None:
            messagebox.showinfo("Cancelled", "ISG status selection was cancelled.")
            return

    # Ask for marker placement
    marker_selector = MarkerSelector(root)
    markers = marker_selector.get_markers()
    if markers is None:
        messagebox.showinfo("Cancelled", "Marker selection was cancelled.")
        return

    # Auto-find 4D motion .ini file and calculate pelvic drop
    pelvic_drop_sentence = None
    motion_ini_path = None
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.ini') and '4d motion' in filename.lower():
            motion_ini_path = os.path.join(folder_path, filename)
            break

    if motion_ini_path:
        print(f"Found motion .ini file: {motion_ini_path}")

        # Parse both .ini files to get beckenhochstand and motion values
        (kyphosis_angle, lordosis_angle, scoliosis_angle,
         surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right, sva_axis, beckenhochstand) = parse_ini_file(ini_path)

        motion_min, motion_max = parse_motion_ini_file(motion_ini_path)

        if beckenhochstand is not None and motion_min is not None and motion_max is not None:
            # Ask user for beckenhochstand side
            beckenhochstand_selector = BeckenhochstandSelector(root)
            beckenhochstand_side = beckenhochstand_selector.get_beckenhochstand_side()

            if beckenhochstand_side is None:
                messagebox.showinfo("Cancelled", "Beckenhochstand selection was cancelled.")
                return

            # Calculate pelvic drop sentence
            pelvic_drop_sentence = calculate_pelvic_drop_sentence(
                beckenhochstand, beckenhochstand_side, motion_min, motion_max
            )
            print(f"Pelvic drop sentence: {pelvic_drop_sentence}")
        else:
            print("Warning: Could not extract all values for pelvic drop calculation")
    else:
        print("Warning: 4D motion .ini file not found in folder")

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
        create_report(patient_full_title, patient_name, patient_dob, report_creator, odt_path, gender, ini_path,
                      sim_performed, isg_right, isg_left, markers, logo_path, second_logo_path,
                      screenshot_path, statik_screenshots, pelvic_drop_sentence, gehen_screenshots, kraft_screenshots)
        messagebox.showinfo("Success", f"Successfully created {odt_path}")

        # Clean up temporary screenshot files
        if screenshot_path and os.path.exists(screenshot_path):
            try:
                os.remove(screenshot_path)
                print(f"Cleaned up temporary file: {screenshot_path}")
            except:
                pass

        # Clean up statik screenshot files
        for i, screenshot_file in enumerate(statik_screenshots, 1):
            if screenshot_file and os.path.exists(screenshot_file):
                try:
                    os.remove(screenshot_file)
                    print(f"Cleaned up statik screenshot {i}: {screenshot_file}")
                except:
                    pass

        # Clean up gehen screenshot files
        for i, screenshot_file in enumerate(gehen_screenshots, 1):
            if screenshot_file and os.path.exists(screenshot_file):
                try:
                    os.remove(screenshot_file)
                    print(f"Cleaned up gehen screenshot {i}: {screenshot_file}")
                except:
                    pass

        # Clean up kraft screenshot files
        for i, screenshot_file in enumerate(kraft_screenshots, 1):
            if screenshot_file and os.path.exists(screenshot_file):
                try:
                    os.remove(screenshot_file)
                    print(f"Cleaned up kraft screenshot {i}: {screenshot_file}")
                except:
                    pass
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Main application
root = tk.Tk()
root.title("Report Generator")

tk.Label(root, text="Generate a new report.").pack(pady=10)
tk.Button(root, text="Generate Report", command=generate_report).pack(pady=10)
tk.Button(root, text="Find Statik.pdf Coordinates", command=find_statik_coordinates).pack(pady=10)
tk.Button(root, text="Find Gehen.pdf Coordinates", command=find_gehen_coordinates).pack(pady=10)
tk.Button(root, text="Find Kraft.pdf Coordinates", command=find_kraft_coordinates).pack(pady=10)

root.mainloop()