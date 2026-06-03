import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import subprocess
import json
import re
from odf.opendocument import OpenDocumentText
from odf.draw import Frame, Image
from odf.text import P, List, ListItem, PageNumber, Span
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.style import Style, TableColumnProperties, ParagraphProperties, TextProperties, ListLevelProperties, PageLayout, PageLayoutProperties, MasterPage, Footer, FooterStyle, TabStops, TabStop, HeaderFooterProperties
from odf.text import ListStyle, ListLevelStyleBullet
from datetime import datetime
from PIL import Image as PILImage, ImageTk
from pdf2image import convert_from_path
import tempfile
import PyPDF2


# Config file path for persistent settings
CONFIG_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_config.json")


def load_config():
    """Load configuration from JSON file"""
    if os.path.exists(CONFIG_FILE_PATH):
        try:
            with open(CONFIG_FILE_PATH, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
    return {}


def save_config(config):
    """Save configuration to JSON file"""
    try:
        with open(CONFIG_FILE_PATH, 'w') as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        print(f"Error saving config: {e}")
        messagebox.showerror("Config Error", f"Could not save configuration: {e}")


def find_libreoffice():
    """Find LibreOffice executable path, especially for Windows."""
    import platform
    import shutil

    # On non-Windows, try the standard command first
    if platform.system() != "Windows":
        if shutil.which("libreoffice"):
            return "libreoffice"
        if shutil.which("soffice"):
            return "soffice"
        return "libreoffice"  # Fall back, let it fail with clear error

    # On Windows, check common installation paths
    possible_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        os.path.expandvars(r"%PROGRAMFILES%\LibreOffice\program\soffice.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\LibreOffice\program\soffice.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    # Check if soffice or libreoffice is in PATH
    if shutil.which("soffice"):
        return shutil.which("soffice")
    if shutil.which("libreoffice"):
        return shutil.which("libreoffice")

    # Not found - return None to allow caller to show proper error
    return None




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


def crop_vgl_screenshot(pdf_path):
    """
    Crop screenshot from vgl.pdf (leg length examination) using hardcoded coordinates.
    Returns path to temporary cropped image file.
    """
    try:
        # Hardcoded crop coordinates as percentages (from user's coordinate finder)
        LEFT_PCT = 2.17
        TOP_PCT = 15.55
        RIGHT_PCT = 59.09
        BOTTOM_PCT = 90.09

        # Convert first page at 300 DPI for high quality
        print("Converting vgl.pdf to high-quality image (300 DPI)...")
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

        print(f"Cropping vgl.pdf to: ({left}, {top}, {right}, {bottom})")

        # Crop the image
        cropped_image = image.crop((left, top, right, bottom))
        crop_width, crop_height = cropped_image.size
        print(f"Cropped vgl image size: {crop_width}x{crop_height} pixels")

        # Save to temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        cropped_image.save(temp_file.name, 'PNG', quality=95)
        print(f"Saved cropped vgl image to: {temp_file.name}")

        return temp_file.name

    except Exception as e:
        print(f"Error cropping vgl.pdf screenshot: {e}")
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
            {"page": 2, "left": 13.73, "top": 27.32, "right": 70.26, "bottom": 82.19},  # Screenshot 1
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


def crop_kraft_screenshots(pdf_path, strength_test_type="Torso + legs"):
    """
    Crop screenshots from kraft.pdf using coordinates based on strength test type.
    Returns list of paths to temporary cropped image files.
    """
    try:
        if strength_test_type == "Torso + legs + shoulders":
            # 6 screenshots: main1, extra1, flat1, main2, extra2, flat2
            screenshots_config = [
                {"page": 1, "left": 1.37, "top": 20.47, "right": 47.98, "bottom": 78.49},   # Screenshot 1 (rechts-links - main)
                {"page": 1, "left": 47.69, "top": 20.31, "right": 93.73, "bottom": 41.18},  # Screenshot 2 (rechts-links - shoulders extra)
                {"page": 1, "left": 8.09, "top": 81.47, "right": 82.68, "bottom": 89.36},   # Screenshot 3 (rechts-links - bottom flat)
                {"page": 2, "left": 2.62, "top": 18.86, "right": 48.77, "bottom": 75.10},   # Screenshot 4 (antagonist - main)
                {"page": 2, "left": 48.60, "top": 18.94, "right": 93.85, "bottom": 32.47},  # Screenshot 5 (antagonist - shoulders extra)
                {"page": 2, "left": 11.00, "top": 80.74, "right": 78.01, "bottom": 89.20},  # Screenshot 6 (antagonist - bottom flat)
            ]
        elif strength_test_type == "Torso + shoulders":
            # 4 screenshots: main1, flat1, main2, flat2
            screenshots_config = [
                {"page": 1, "left": 1.14, "top": 20.47, "right": 48.43, "bottom": 64.22},   # Screenshot 1 (rechts-links - main)
                {"page": 1, "left": 8.09, "top": 81.47, "right": 82.68, "bottom": 89.36},   # Screenshot 2 (rechts-links - bottom flat)
                {"page": 2, "left": 1.14, "top": 20.47, "right": 48.43, "bottom": 64.22},   # Screenshot 3 (antagonist - main)
                {"page": 2, "left": 11.00, "top": 80.74, "right": 78.01, "bottom": 89.20},  # Screenshot 4 (antagonist - bottom flat)
            ]
        elif strength_test_type == "Legs + shoulders":
            # 4 screenshots: main1, flat1, main2, flat2 (standard two-page layout)
            screenshots_config = [
                {"page": 1, "left": 1.11, "top": 21.11, "right": 48.83, "bottom": 78.97},   # Screenshot 1 (rechts-links - main)
                {"page": 1, "left": 8.09, "top": 81.47, "right": 82.68, "bottom": 89.36},   # Screenshot 2 (rechts-links - bottom flat)
                {"page": 2, "left": 1.65, "top": 18.37, "right": 49.46, "bottom": 55.92},   # Screenshot 3 (antagonist - main)
                {"page": 2, "left": 11.00, "top": 80.74, "right": 78.01, "bottom": 89.20},  # Screenshot 4 (antagonist - bottom flat)
            ]
        elif strength_test_type == "Legs":
            # 4 screenshots: main1 (smaller), flat1, main2 (smaller), flat2
            # Smaller main screenshots so both sections fit on one page
            screenshots_config = [
                {"page": 1, "left": 0.68, "top": 20.87, "right": 49.46, "bottom": 57.45},   # Screenshot 1 (rechts-links - main)
                {"page": 1, "left": 8.09, "top": 81.47, "right": 82.68, "bottom": 89.36},   # Screenshot 2 (rechts-links - bottom flat)
                {"page": 2, "left": 1.77, "top": 19.10, "right": 49.00, "bottom": 42.47},   # Screenshot 3 (antagonist - main)
                {"page": 2, "left": 11.00, "top": 80.74, "right": 78.01, "bottom": 89.20},  # Screenshot 4 (antagonist - bottom flat)
            ]
        else:
            # Default: Torso + legs - 4 screenshots (original coordinates)
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


def crop_laufen_screenshots(hp_pdf_path, ios_pdf_path):
    """
    Crop 8 screenshots from hp.pdf and ios.pdf using hardcoded coordinates.
    Returns list of paths to temporary cropped image files.
    Screenshots 1-7 from hp.pdf, Screenshot 8 from ios.pdf
    """
    try:
        # Hardcoded crop coordinates for 8 screenshots
        # Screenshots 1-7 from hp.pdf, Screenshot 8 from ios.pdf
        screenshots_config = [
            {"pdf": hp_pdf_path, "page": 1, "left": 33.45, "top": 15.71, "right": 92.08, "bottom": 90.73},  # Screenshot 1
            {"pdf": hp_pdf_path, "page": 2, "left": 0.40, "top": 16.36, "right": 37.09, "bottom": 84.93},   # Screenshot 2
            {"pdf": hp_pdf_path, "page": 3, "left": 14.02, "top": 18.69, "right": 92.14, "bottom": 77.20},  # Screenshot 3
            {"pdf": hp_pdf_path, "page": 6, "left": 11.97, "top": 18.45, "right": 91.97, "bottom": 81.55},  # Screenshot 4
            {"pdf": hp_pdf_path, "page": 4, "left": 1.48, "top": 15.47, "right": 92.93, "bottom": 87.11},   # Screenshot 5
            {"pdf": hp_pdf_path, "page": 5, "left": 0.85, "top": 14.59, "right": 93.73, "bottom": 86.30},   # Screenshot 6
            {"pdf": hp_pdf_path, "page": 8, "left": 3.25, "top": 20.31, "right": 93.39, "bottom": 87.11},   # Screenshot 7
            {"pdf": ios_pdf_path, "page": 1, "left": 1.99, "top": 18.86, "right": 93.68, "bottom": 87.35},  # Screenshot 8
        ]

        screenshot_paths = []

        for i, config in enumerate(screenshots_config, 1):
            print(f"\nProcessing Laufen Screenshot {i} from {config['pdf']} page {config['page']}...")

            # Convert specific page at 300 DPI for high quality
            images = convert_from_path(config['pdf'], dpi=300, first_page=config['page'], last_page=config['page'])
            if not images:
                print(f"Warning: Could not convert page {config['page']} from {config['pdf']}")
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
            print(f"Laufen Screenshot {i} cropped: {crop_width}x{crop_height} pixels")

            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            cropped_image.save(temp_file.name, 'PNG', quality=95)
            screenshot_paths.append(temp_file.name)

        return screenshot_paths

    except Exception as e:
        print(f"Error cropping laufen screenshots: {e}")
        import traceback
        traceback.print_exc()
        return []


def crop_ios_pedografie_screenshots(pdf_path):
    """
    Crop 2 screenshots from ios.pdf for Dynamische Pedografie section.
    Uses same coordinates as gehen screenshots 7 and 8 (pages 7 and 8).
    Returns list of paths to temporary cropped image files.
    """
    try:
        # Same coordinates as gehen screenshots 7 and 8, but from pages 1 and 2 of ios.pdf
        screenshots_config = [
            {"page": 1, "left": 3.13, "top": 19.10, "right": 94.02, "bottom": 87.11},   # Screenshot 1 (same coords as gehen screenshot 7)
            {"page": 2, "left": 14.36, "top": 18.69, "right": 93.90, "bottom": 84.05},  # Screenshot 2 (same coords as gehen screenshot 8)
        ]

        screenshot_paths = []

        for i, config in enumerate(screenshots_config, 1):
            print(f"\nProcessing IOS Pedografie Screenshot {i} from page {config['page']}...")

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
            print(f"IOS Pedografie Screenshot {i} cropped: {crop_width}x{crop_height} pixels")

            # Save to temporary file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            cropped_image.save(temp_file.name, 'PNG', quality=95)
            screenshot_paths.append(temp_file.name)

        return screenshot_paths

    except Exception as e:
        print(f"Error cropping ios pedografie screenshots: {e}")
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
            {"page": 3, "left": 15.27, "top": 17.49, "right": 92.72, "bottom": 76.87},  # Screenshot 3 (Ganganalyse - page 3)
            {"page": 5, "left": 13.73, "top": 18.69, "right": 92.08, "bottom": 80.82},  # Screenshot 4 (Ganganalyse - page 5)
            {"page": 4, "left": 2.62, "top": 17.57, "right": 92.88, "bottom": 87.11},   # Screenshot 5 (Ganganalyse - page 4)
            {"page": 6, "left": 1.82, "top": 22.24, "right": 93.56, "bottom": 96.37},   # Screenshot 6 (Ganganalyse - page 6)
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


def extract_patient_info_from_pdf(pdf_path):
    """
    Extract patient name, date of birth, and measurement date from PDF.
    Returns (patient_name, patient_dob, measurement_date) as strings.
    Patient name is returned as "Last name, First name" format.
    """
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if len(reader.pages) == 0:
                return None, None, None

            # Get text from first page
            page = reader.pages[0]
            text = page.extract_text()

            if not text:
                return None, None, None

            patient_name = None
            patient_dob = None
            measurement_date = None

            # Extract name and DOB from "Name: First Last (* DD.MM.YYYY)"
            name_pattern = r"Name:\s*([^(]+)\s*\(\*\s*(\d{2}\.\d{2}\.\d{4})\)"
            name_match = re.search(name_pattern, text)
            if name_match:
                full_name = name_match.group(1).strip()
                patient_dob = name_match.group(2).strip()

                # Convert "First Last" to "Last, First"
                name_parts = full_name.split()
                if len(name_parts) >= 2:
                    first_name = name_parts[0]
                    last_name = " ".join(name_parts[1:])
                    patient_name = f"{last_name}, {first_name}"
                else:
                    patient_name = full_name

            # Extract measurement date from "4D average vom DD.MM.YYYY"
            date_pattern = r"4D average vom\s*(\d{2}\.\d{2}\.\d{4})"
            date_match = re.search(date_pattern, text, re.IGNORECASE)
            if date_match:
                measurement_date = date_match.group(1).strip()

            return patient_name, patient_dob, measurement_date

    except Exception as e:
        print(f"Error extracting patient info from PDF: {e}")
        return None, None, None


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
    """Parse the 4D motion .ini file and extract pelvic drop mean, min and max values"""
    try:
        with open(ini_path, 'r', encoding='utf-16') as f:
            lines = f.readlines()

        pelvic_drop_mean = None
        pelvic_drop_min = None
        pelvic_drop_max = None

        # Line 11 (index 10) for pelvic drop mean (column 2), min (column 3) and max (column 4)
        if len(lines) > 10:
            parts = lines[10].split('\t')
            # Column 2 (index 1) for mean
            if len(parts) >= 2:
                value_str = parts[1].strip().replace(',', '.')
                try:
                    pelvic_drop_mean = float(value_str)
                except ValueError:
                    pass
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

        return pelvic_drop_mean, pelvic_drop_min, pelvic_drop_max
    except Exception as e:
        print(f"Error parsing motion INI file: {e}")
        return None, None, None


def calculate_pelvic_drop_sentence(beckenhochstand, motion_mean, motion_min, motion_max):
    """Calculate pelvic drop values and generate appropriate sentence

    Args:
        beckenhochstand: Value from 4D average .ini file (X)
        motion_mean: Mean value from 4D motion .ini file (p_mean)
        motion_min: Min value from 4D motion .ini file (p_min)
        motion_max: Max value from 4D motion .ini file (p_max)

    Returns:
        Sentence describing the pelvic drop
    """
    if motion_mean is None or motion_min is None or motion_max is None:
        return None

    # Use absolute value of X for calculation
    abs_x = abs(beckenhochstand) if beckenhochstand is not None else 0

    # Calculate absolute differences
    abs_diff_min = abs(motion_mean - motion_min)
    abs_diff_max = abs(motion_mean - motion_max)

    # Calculate A and B based on X sign
    if beckenhochstand is not None and beckenhochstand >= 0:
        # X is positive
        A = motion_mean - abs_x - abs_diff_min
        B = motion_mean - abs_x + abs_diff_max
    else:
        # X is negative or None
        A = motion_mean + abs_x - abs_diff_min
        B = motion_mean + abs_x + abs_diff_max

    # Take absolute values for display
    A = abs(A)
    B = abs(B)

    # Calculate absolute difference
    abs_difference = abs(A - B)

    # Format values with one decimal place
    A_str = f"{A:.1f}"
    B_str = f"{B:.1f}"

    # Generate sentence based on difference (B first, then A)
    if abs_difference <= 2:
        # Symmetric pelvic drop
        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet ein symmetrischer Pelvic Drop von ({B_str}°L | {A_str}°R) statt"
    else:
        # Asymmetric pelvic drop - determine which stance phase
        if B > A:
            stance_phase = "rechten"
        else:
            stance_phase = "linken"

        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet während der {stance_phase} Standphase ein asymmetrischer Pelvic Drop von ({B_str}°L | {A_str}°R) statt"

    return sentence


def generate_pelvic_drop_sentence_from_custom(r_value, l_value):
    """Generate pelvic drop sentence from custom R and L values

    Args:
        r_value: Custom R (Rechts) value
        l_value: Custom L (Links) value

    Returns:
        Sentence describing the pelvic drop
    """
    # Calculate absolute difference
    abs_difference = abs(r_value - l_value)

    # Format values with one decimal place
    r_str = f"{r_value:.1f}"
    l_str = f"{l_value:.1f}"

    # Generate sentence based on difference (L first, then R)
    if abs_difference <= 2:
        # Symmetric pelvic drop
        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet ein symmetrischer Pelvic Drop von ({l_str}°L | {r_str}°R) statt"
    else:
        # Asymmetric pelvic drop - determine which stance phase
        if l_value > r_value:
            stance_phase = "rechten"
        else:
            stance_phase = "linken"

        sentence = f"Unter Berücksichtigung der geklebten Beckenmarker findet während der {stance_phase} Standphase ein asymmetrischer Pelvic Drop von ({l_str}°L | {r_str}°R) statt"

    return sentence


def classify_kyphosis(angle):
    """Classify kyphosis angle (same for both genders)

    Boundary values (X.0) go to the 'healthier' classification:
    - Normal range includes its boundaries (39 <= angle <= 56)
    - Values 0.1 beyond boundary go to the worse category
    """
    if angle < 31:
        return "Eine Hypokyphose"
    elif 31 <= angle < 39:
        return "Eine Tendenz zur Hypokyphose"
    elif 39 <= angle <= 56:
        return "Ein normgerechter Kyphosewinkel"
    elif 56 < angle <= 64:
        return "Eine Tendenz zur Hyperkyphose"
    else:  # > 64
        return "Eine Hyperkyphose"


def classify_lordosis(angle, gender):
    """Classify lordosis angle based on gender

    Boundary values (X.0) go to the 'healthier' classification:
    - Normal range includes its boundaries
    - Values 0.1 beyond boundary go to the worse category
    """
    if gender == "Female":
        if angle < 24:
            return "Eine Hypolordose"
        elif 24 <= angle < 33:
            return "Eine Tendenz zur Hypolordose"
        elif 33 <= angle <= 51:
            return "Ein normgerechter Lordosewinkel"
        elif 51 < angle <= 60:
            return "Eine Tendenz zur Hyperlordose"
        else:  # > 60
            return "Eine Hyperlordose"
    else:  # Male
        if angle < 18:
            return "Eine Hypolordose"
        elif 18 <= angle < 26:
            return "Eine Tendenz zur Hypolordose"
        elif 26 <= angle <= 42:
            return "Ein normgerechter Lordosewinkel"
        elif 42 < angle <= 49:
            return "Eine Tendenz zur Hyperlordose"
        else:  # > 49
            return "Eine Hyperlordose"


def classify_scoliosis(angle, surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right):
    """Classify scoliosis angle and add surface rotation and lateral deviation"""
    # Base classification
    if abs(angle) < 10:
        base_text = "Eine skoliotische Haltungsstörung"
    else:
        base_text = "Eine Skoliose"

    # Determine surface rotation direction and value
    abs_left_rot = abs(surf_rot_left) if surf_rot_left is not None else 0
    abs_right_rot = abs(surf_rot_right) if surf_rot_right is not None else 0

    if abs_left_rot > abs_right_rot:
        rotation_direction = "rechts"
        rotation_value = abs_left_rot
    else:
        rotation_direction = "links"
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
                  screenshot_path=None, statik_screenshots=None, pelvic_drop_sentence=None, gehen_screenshots=None, kraft_screenshots=None,
                  beinachsen_texts=None, ganganalyse_texts=None, therapie_texts=None, measurement_type="Gehen", ios_pedografie_screenshots=None,
                  measurement_date=None, leg_length_texts=None, vgl_screenshot=None, strength_test_type="Torso + legs"):
    doc = OpenDocumentText()

    # Create page layout for first page (no footer)
    first_page_layout = PageLayout(name="FirstPageLayout")
    first_page_layout.addElement(PageLayoutProperties(
        pagewidth="21cm",
        pageheight="29.7cm",
        marginleft="2cm",
        marginright="2cm",
        margintop="2cm",
        marginbottom="2cm"
    ))
    doc.automaticstyles.addElement(first_page_layout)

    # Create page layout for standard pages (with footer)
    # Reduced bottom margin to compensate for footer content
    standard_page_layout = PageLayout(name="StandardPageLayout")
    standard_page_layout.addElement(PageLayoutProperties(
        pagewidth="21cm",
        pageheight="29.7cm",
        marginleft="2cm",
        marginright="2cm",
        margintop="2cm",
        marginbottom="0.5cm"
    ))
    # Add footer style
    footer_style = FooterStyle()
    footer_style.addElement(HeaderFooterProperties(minheight="0cm"))
    standard_page_layout.addElement(footer_style)
    doc.automaticstyles.addElement(standard_page_layout)

    # Create footer paragraph style (italic, 9pt, left aligned for address)
    footer_p_style = Style(name="FooterParagraphStyle", family="paragraph")
    footer_p_style.addElement(ParagraphProperties(textalign="left"))
    footer_p_style.addElement(TextProperties(fontsize="9pt", fontstyle="italic"))
    doc.styles.addElement(footer_p_style)

    # Create footer paragraph style for page number (right aligned, 9pt)
    footer_page_num_style = Style(name="FooterPageNumStyle", family="paragraph")
    footer_page_num_style.addElement(ParagraphProperties(textalign="right"))
    footer_page_num_style.addElement(TextProperties(fontsize="9pt"))
    doc.styles.addElement(footer_page_num_style)

    # Create first page master (no footer) - must be listed first to be the default
    first_master = MasterPage(name="First", pagelayoutname="FirstPageLayout")
    doc.masterstyles.addElement(first_master)

    # Create standard master page (with footer and page numbers)
    standard_master = MasterPage(name="Standard", pagelayoutname="StandardPageLayout")
    # Add footer content: address on left, page number on right (separate lines)
    footer = Footer()
    # Page number line (right aligned, at top of footer) - subtract 1 so page 2 shows as "1"
    footer_pnum = P(stylename="FooterPageNumStyle")
    footer_pnum.addElement(PageNumber(numformat="1", pageadjust="-1"))
    footer.addElement(footer_pnum)
    # Address lines (left aligned, italic)
    footer_p1 = P(stylename="FooterParagraphStyle")
    footer_p1.addText("Orthopassion - Privatpraxis für regenerative Orthopädie und Osteopathie")
    footer.addElement(footer_p1)
    footer_p2 = P(stylename="FooterParagraphStyle")
    footer_p2.addText("Gartenstraße 28, 79098 Freiburg im Breisgau")
    footer.addElement(footer_p2)
    footer_p3 = P(stylename="FooterParagraphStyle")
    footer_p3.addText("0761 – 769 911 66, Info@orthopassion.de")
    footer.addElement(footer_p3)
    standard_master.addElement(footer)
    doc.masterstyles.addElement(standard_master)

    # Style for first page content (uses First master - no footer)
    first_page_content_style = Style(name="FirstPageContentStyle", family="paragraph", masterpagename="First")
    doc.styles.addElement(first_page_content_style)

    # Define style for centered paragraphs
    center_p_style = Style(name="CenterParagraph", family="paragraph")
    center_p_style.addElement(ParagraphProperties(textalign="center"))
    doc.styles.addElement(center_p_style)

    # Define style for first page title (18pt, bold, centered)
    title_style = Style(name="TitleStyle", family="paragraph")
    title_style.addElement(ParagraphProperties(textalign="center"))
    title_style.addElement(TextProperties(fontsize="18pt", fontweight="bold"))
    doc.styles.addElement(title_style)

    # Define style for first page patient info (14pt, not bold, centered)
    patient_info_style = Style(name="PatientInfoStyle", family="paragraph")
    patient_info_style.addElement(ParagraphProperties(textalign="center"))
    patient_info_style.addElement(TextProperties(fontsize="14pt"))
    doc.styles.addElement(patient_info_style)

    # Define style for headings with page break (bold, 14pt) - uses Standard master with footer, resets page to 1
    heading_with_break_style = Style(name="HeadingWithBreakStyle", family="paragraph", masterpagename="Standard")
    hwb_props = ParagraphProperties(breakbefore="page")
    hwb_props.setAttrNS("urn:oasis:names:tc:opendocument:xmlns:style:1.0", "style:page-number", "1")
    heading_with_break_style.addElement(hwb_props)
    heading_with_break_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
    doc.styles.addElement(heading_with_break_style)

    # Define style for subsequent page breaks (bold, 14pt) - continues on Standard master
    heading_page_break_style = Style(name="HeadingPageBreakStyle", family="paragraph")
    heading_page_break_style.addElement(ParagraphProperties(breakbefore="page"))
    heading_page_break_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
    doc.styles.addElement(heading_page_break_style)

    # Define style for headings (bold, 14pt)
    heading_style = Style(name="HeadingStyle", family="paragraph")
    heading_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
    doc.styles.addElement(heading_style)

    # Define style for small headings (bold, 12pt)
    small_heading_style = Style(name="SmallHeadingStyle", family="paragraph")
    small_heading_style.addElement(TextProperties(fontsize="12pt", fontweight="bold"))
    doc.styles.addElement(small_heading_style)

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
    bullet_level.addElement(ListLevelProperties(spacebefore="0.5cm", minlabelwidth="0.5cm"))
    bullet_list_style.addElement(bullet_level)
    doc.styles.addElement(bullet_list_style)

    # Define style for indented bullet text (12pt font, 1.5 line spacing, with left margin)
    indented_bullet_text_style = Style(name="IndentedBulletTextStyle", family="paragraph")
    indented_bullet_para_props = ParagraphProperties(lineheight="150%", marginleft="1cm")
    indented_bullet_text_style.addElement(indented_bullet_para_props)
    indented_bullet_text_style.addElement(TextProperties(fontsize="12pt"))
    doc.styles.addElement(indented_bullet_text_style)

    # Define style for page break (continues page numbering)
    page_break_style = Style(name="PageBreakStyle", family="paragraph")
    page_break_style.addElement(ParagraphProperties(breakbefore="page"))
    doc.styles.addElement(page_break_style)

    # Define style for first page content (uses First master - no footer)
    first_page_style = Style(name="FirstPageStyle", family="paragraph", masterpagename="First")
    doc.styles.addElement(first_page_style)

    # Add invisible paragraph to establish First master page (no footer on page 1)
    doc.text.addElement(P(text="", stylename="FirstPageContentStyle"))

    # First page centered text content
    today = datetime.now().strftime("%d.%m.%Y")
    # Use measurement_date for headline, fall back to today if not provided
    headline_date = measurement_date if measurement_date else today

    # Title: 18pt, bold, centered
    headline = P(text=f"Befundbericht zur Bewegungsanalyse vom {headline_date}", stylename="TitleStyle")
    doc.text.addElement(headline)

    # Empty line
    doc.text.addElement(P(text=""))

    # Patient info: 14pt, not bold, centered
    patient_info = P(text=f"{patient_full_title} {patient_name}, geb. am: {patient_dob}", stylename="PatientInfoStyle")
    doc.text.addElement(patient_info)

    # Empty line
    doc.text.addElement(P(text=""))

    # Creator info: 14pt, not bold, centered
    creator_info = P(text=f"Bericht erstellt von: {report_creator} am {today}", stylename="PatientInfoStyle")
    doc.text.addElement(creator_info)

    # Add spacing before second logo to center it vertically
    for _ in range(6):
        doc.text.addElement(P(text=""))

    # Second Logo (Centered, vertically centered on page)
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

    # Add spacing before bottom logos to push them closer to page bottom
    for _ in range(9):
        doc.text.addElement(P(text=""))

    # Bottom logos - first logo (logo_orthopassion.png) and new logo (logo_4_orthopassion.png) side by side
    # Create a table for side-by-side layout
    bottom_logos_table = Table(name="BottomLogosTable")

    # Define two equal columns for the logos
    bottom_logo_col_style_left = Style(name="BottomLogoColumnStyleLeft", family="table-column")
    bottom_logo_col_style_left.addElement(TableColumnProperties(columnwidth="9cm"))
    doc.styles.addElement(bottom_logo_col_style_left)

    bottom_logo_col_style_right = Style(name="BottomLogoColumnStyleRight", family="table-column")
    bottom_logo_col_style_right.addElement(TableColumnProperties(columnwidth="9cm"))
    doc.styles.addElement(bottom_logo_col_style_right)

    bottom_logos_table.addElement(TableColumn(stylename="BottomLogoColumnStyleLeft"))
    bottom_logos_table.addElement(TableColumn(stylename="BottomLogoColumnStyleRight"))

    bottom_logos_row = TableRow()

    # Left cell - first logo (logo_orthopassion.png) at 8cm wide
    left_logo_cell = TableCell()
    if logo_path and os.path.exists(logo_path):
        try:
            with PILImage.open(logo_path) as logo_img:
                original_logo_width_px, original_logo_height_px = logo_img.size

            logo_width_cm = 8.0
            if original_logo_width_px > 0:
                aspect_ratio = original_logo_height_px / original_logo_width_px
            else:
                aspect_ratio = 1

            logo_height_cm = logo_width_cm * aspect_ratio

            logo_p = P()
            logo_frame = Frame(
                name="BottomLogoFrame1",
                width=f"{logo_width_cm}cm",
                height=f"{logo_height_cm}cm",
                anchortype="paragraph"
            )
            logo_href = doc.addPicture(logo_path)
            if logo_href:
                logo_frame.addElement(Image(href=logo_href))
                logo_p.addElement(logo_frame)
                left_logo_cell.addElement(logo_p)
            else:
                left_logo_cell.addElement(P(text=f"Logo error: {logo_path}"))
                messagebox.showwarning("Logo Error", f"Could not embed logo from {logo_path}")
        except FileNotFoundError:
            left_logo_cell.addElement(P(text=f"Logo file not found: {logo_path}"))
            messagebox.showerror("Logo Error", f"Logo file not found at: {logo_path}")
        except Exception as e:
            left_logo_cell.addElement(P(text=f"Error adding logo: {e}"))
            messagebox.showerror("Logo Error", f"An error occurred while adding the logo: {e}")
    else:
        left_logo_cell.addElement(P(text=""))

    bottom_logos_row.addElement(left_logo_cell)

    # Right cell - new logo (logo_4_orthopassion.png) at 8cm wide
    right_logo_cell = TableCell()
    # Construct the path for logo_4_orthopassion.png
    if logo_path:
        logo_4_path = logo_path.replace("logo_orthopassion.png", "logo_4_orthopassion.png")
    else:
        logo_4_path = None

    if logo_4_path and os.path.exists(logo_4_path):
        try:
            with PILImage.open(logo_4_path) as logo_img:
                original_logo_width_px, original_logo_height_px = logo_img.size

            logo_width_cm = 8.0
            if original_logo_width_px > 0:
                aspect_ratio = original_logo_height_px / original_logo_width_px
            else:
                aspect_ratio = 1

            logo_height_cm = logo_width_cm * aspect_ratio

            logo_p = P()
            logo_frame = Frame(
                name="BottomLogoFrame2",
                width=f"{logo_width_cm}cm",
                height=f"{logo_height_cm}cm",
                anchortype="paragraph"
            )
            logo_href = doc.addPicture(logo_4_path)
            if logo_href:
                logo_frame.addElement(Image(href=logo_href))
                logo_p.addElement(logo_frame)
                right_logo_cell.addElement(logo_p)
            else:
                right_logo_cell.addElement(P(text=f"Logo 4 error: {logo_4_path}"))
                messagebox.showwarning("Logo 4 Error", f"Could not embed logo 4 from {logo_4_path}")
        except FileNotFoundError:
            right_logo_cell.addElement(P(text=f"Logo 4 file not found: {logo_4_path}"))
            messagebox.showerror("Logo 4 Error", f"Logo 4 file not found at: {logo_4_path}")
        except Exception as e:
            right_logo_cell.addElement(P(text=f"Error adding logo 4: {e}"))
            messagebox.showerror("Logo 4 Error", f"An error occurred while adding logo 4: {e}")
    else:
        right_logo_cell.addElement(P(text=""))

    bottom_logos_row.addElement(right_logo_cell)
    bottom_logos_table.addElement(bottom_logos_row)
    doc.text.addElement(bottom_logos_table)

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
    doc.text.addElement(P(text="Statische Analyse", stylename="HeadingPageBreakStyle"))
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

        # Add leg length examination page (conditional)
        if leg_length_texts and vgl_screenshot and os.path.exists(vgl_screenshot):
            # Page break with heading
            doc.text.addElement(P(text="Messergebnis Beinlängendifferenz", stylename="HeadingPageBreakStyle"))

            # One empty line
            doc.text.addElement(P(text=""))

            # Add bullet points from user input
            if len(leg_length_texts) > 0:
                leg_length_bullet_list = List(stylename="BulletList")
                for text in leg_length_texts:
                    bullet_item = ListItem()
                    bullet_item.addElement(P(text=text, stylename="BulletTextStyle"))
                    leg_length_bullet_list.addElement(bullet_item)
                doc.text.addElement(leg_length_bullet_list)

            # 3 empty lines
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))

            # Add vgl screenshot (centered)
            try:
                with PILImage.open(vgl_screenshot) as vgl_img:
                    vgl_width_px, vgl_height_px = vgl_img.size

                # Set image width to 16cm and calculate height maintaining aspect ratio
                vgl_frame_width_cm = 16.0
                if vgl_width_px > 0:
                    vgl_aspect_ratio = vgl_height_px / vgl_width_px
                else:
                    vgl_aspect_ratio = 1

                vgl_frame_height_cm = vgl_frame_width_cm * vgl_aspect_ratio

                # Create centered paragraph for vgl screenshot
                centered_p_vgl = P(stylename="CenterParagraph")
                vgl_frame = Frame(
                    name="VglScreenshotFrame",
                    width=f"{vgl_frame_width_cm}cm",
                    height=f"{vgl_frame_height_cm}cm",
                    anchortype="paragraph"
                )
                vgl_href = doc.addPicture(vgl_screenshot)
                if vgl_href:
                    vgl_frame.addElement(Image(href=vgl_href))
                    centered_p_vgl.addElement(vgl_frame)
                    doc.text.addElement(centered_p_vgl)
                    print(f"Vgl screenshot inserted: {vgl_frame_width_cm}cm x {vgl_frame_height_cm}cm")
                else:
                    print(f"Warning: Could not embed vgl screenshot from {vgl_screenshot}")

            except Exception as e:
                print(f"Error adding vgl screenshot: {e}")
                import traceback
                traceback.print_exc()

        # Add new section: Statische Beinachsen- und Haltungsanalyse
        if statik_screenshots and len(statik_screenshots) >= 7:
            # Page break before new section
            doc.text.addElement(P(text="Statische Beinachsen- und Haltungsanalyse", stylename="HeadingPageBreakStyle"))

            # Add bullet points from user input
            if beinachsen_texts and len(beinachsen_texts) > 0:
                statik_bullet_list = List(stylename="BulletList")
                for text in beinachsen_texts:
                    bullet_item = ListItem()
                    bullet_item.addElement(P(text=text, stylename="BulletTextStyle"))
                    statik_bullet_list.addElement(bullet_item)
                doc.text.addElement(statik_bullet_list)

            # Add Screenshots 1 and 2 (16cm width, centered) - at the bottom
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
            doc.text.addElement(P(text="Statische Pedobarografie", stylename="HeadingPageBreakStyle"))

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

        # Add new section: Dynamische Beckenanalyse (only for Gehen/Laufen)
        if measurement_type in ["Gehen", "Laufen"] and pelvic_drop_sentence:
            # Page break before new section
            doc.text.addElement(P(text="Dynamische Beckenanalyse", stylename="HeadingPageBreakStyle"))
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

        # Add new section: Dynamische Wirbelsäulenanalyse with gehen screenshot 2 (only for Gehen/Laufen)
        if measurement_type in ["Gehen", "Laufen"] and gehen_screenshots and len(gehen_screenshots) >= 2 and gehen_screenshots[1] and os.path.exists(gehen_screenshots[1]):
            try:
                # Page break with heading
                doc.text.addElement(P(text="Dynamische Wirbelsäulenanalyse", stylename="HeadingPageBreakStyle"))

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

        # Add new section: Ganganalyse (only for Gehen/Laufen)
        if measurement_type in ["Gehen", "Laufen"] and gehen_screenshots and len(gehen_screenshots) >= 6:
            # Page break with heading - use "Laufanalyse" for Laufen, "Ganganalyse" for Gehen
            section_title = "Laufanalyse" if measurement_type == "Laufen" else "Ganganalyse"
            doc.text.addElement(P(text=section_title, stylename="HeadingPageBreakStyle"))
            doc.text.addElement(P(text=""))

            # Add bullet points from user input
            if ganganalyse_texts and len(ganganalyse_texts) > 0:
                ganganalyse_bullet_list = List(stylename="BulletList")
                for text in ganganalyse_texts:
                    bullet_item = ListItem()
                    bullet_item.addElement(P(text=text, stylename="BulletTextStyle"))
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
        # For Gehen/Laufen: use gehen_screenshots (indices 6 and 7)
        # For Statik/IOS: use ios_pedografie_screenshots (indices 0 and 1)
        if measurement_type in ["Gehen", "Laufen"] and gehen_screenshots and len(gehen_screenshots) >= 8:
            # Page break with heading
            doc.text.addElement(P(text="Dynamische Pedografie", stylename="HeadingPageBreakStyle"))

            # Add 2 empty lines after headline
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))

            # Add "Laufen" headline if measurement type is Laufen
            if measurement_type == "Laufen":
                doc.text.addElement(P(text="Laufen", stylename="SmallHeadingStyle"))
                doc.text.addElement(P(text=""))

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

            # Add "Gehen" headline if measurement type is Laufen
            if measurement_type == "Laufen":
                doc.text.addElement(P(text="Gehen", stylename="SmallHeadingStyle"))
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

        elif measurement_type in ["Statik", "IOS"] and ios_pedografie_screenshots and len(ios_pedografie_screenshots) >= 2:
            # Page break with heading
            doc.text.addElement(P(text="Dynamische Pedografie", stylename="HeadingPageBreakStyle"))

            # Add 2 empty lines after headline
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))

            # Add Screenshot 1 (index 0) from ios.pdf
            if ios_pedografie_screenshots[0] and os.path.exists(ios_pedografie_screenshots[0]):
                try:
                    with PILImage.open(ios_pedografie_screenshots[0]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="IOSPedografieScreenshot1",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(ios_pedografie_screenshots[0])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"IOS Pedografie Screenshot 1 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding ios pedografie screenshot 1: {e}")

            # Add 2 empty lines spacing
            doc.text.addElement(P(text=""))
            doc.text.addElement(P(text=""))

            # Add Screenshot 2 (index 1) from ios.pdf
            if ios_pedografie_screenshots[1] and os.path.exists(ios_pedografie_screenshots[1]):
                try:
                    with PILImage.open(ios_pedografie_screenshots[1]) as img:
                        img_width_px, img_height_px = img.size

                    frame_width_cm = 16.0
                    aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                    frame_height_cm = frame_width_cm * aspect_ratio

                    centered_p = P(stylename="CenterParagraph")
                    frame = Frame(
                        name="IOSPedografieScreenshot2",
                        width=f"{frame_width_cm}cm",
                        height=f"{frame_height_cm}cm",
                        anchortype="paragraph"
                    )
                    href = doc.addPicture(ios_pedografie_screenshots[1])
                    if href:
                        frame.addElement(Image(href=href))
                        centered_p.addElement(frame)
                        doc.text.addElement(centered_p)
                        print(f"IOS Pedografie Screenshot 2 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                except Exception as e:
                    print(f"Error adding ios pedografie screenshot 2: {e}")

        # Add new section: Kraftanalyse Vergleich rechts - links
        # Layout depends on strength_test_type:
        # "Torso + legs" / "Torso + shoulders": indices [0]=main, [1]=flat, [2]=main, [3]=flat
        # "Torso + legs + shoulders": indices [0]=main, [1]=extra, [2]=flat, [3]=main, [4]=extra, [5]=flat
        if strength_test_type == "Torso + legs + shoulders":
            # Page 1: rechts-links with 3 screenshots (main + extra + flat)
            if kraft_screenshots and len(kraft_screenshots) >= 3:
                doc.text.addElement(P(text="Kraftanalyse Vergleich rechts - links", stylename="HeadingPageBreakStyle"))
                doc.text.addElement(P(text=""))

                # Main screenshot (index 0)
                if kraft_screenshots[0] and os.path.exists(kraft_screenshots[0]):
                    try:
                        with PILImage.open(kraft_screenshots[0]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot1", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[0])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 1 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 1: {e}")

                # Extra shoulders screenshot (index 1) - directly underneath
                if kraft_screenshots[1] and os.path.exists(kraft_screenshots[1]):
                    try:
                        with PILImage.open(kraft_screenshots[1]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot2", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[1])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 2 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 2: {e}")

                # Flat bottom screenshot (index 2) - no spacing, fits directly under extra
                if kraft_screenshots[2] and os.path.exists(kraft_screenshots[2]):
                    try:
                        with PILImage.open(kraft_screenshots[2]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot3", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[2])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 3 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 3: {e}")

            # Page 2: antagonist-agonist with 3 screenshots (main + extra + flat)
            if kraft_screenshots and len(kraft_screenshots) >= 6:
                doc.text.addElement(P(text="Kraftanalyse Vergleich Antagonist - Agonist", stylename="HeadingPageBreakStyle"))
                doc.text.addElement(P(text=""))

                # Main screenshot (index 3)
                if kraft_screenshots[3] and os.path.exists(kraft_screenshots[3]):
                    try:
                        with PILImage.open(kraft_screenshots[3]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot4", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[3])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 4 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 4: {e}")

                # Extra shoulders screenshot (index 4) - directly underneath
                if kraft_screenshots[4] and os.path.exists(kraft_screenshots[4]):
                    try:
                        with PILImage.open(kraft_screenshots[4]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot5", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[4])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 5 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 5: {e}")

                # Flat bottom screenshot (index 5) - no spacing, fits directly under extra
                if kraft_screenshots[5] and os.path.exists(kraft_screenshots[5]):
                    try:
                        with PILImage.open(kraft_screenshots[5]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot6", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[5])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 6 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 6: {e}")

        else:
            # "Torso + legs" or "Torso + shoulders": 4 screenshots [main, flat, main, flat]
            if kraft_screenshots and len(kraft_screenshots) >= 2:
                doc.text.addElement(P(text="Kraftanalyse Vergleich rechts - links", stylename="HeadingPageBreakStyle"))
                doc.text.addElement(P(text=""))

                # Main screenshot (index 0)
                if kraft_screenshots[0] and os.path.exists(kraft_screenshots[0]):
                    try:
                        with PILImage.open(kraft_screenshots[0]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot1", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[0])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 1 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 1: {e}")

                # Flat bottom screenshot (index 1)
                if kraft_screenshots[1] and os.path.exists(kraft_screenshots[1]):
                    try:
                        with PILImage.open(kraft_screenshots[1]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot2", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[1])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 2 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 2: {e}")

            # Antagonist-Agonist page
            if kraft_screenshots and len(kraft_screenshots) >= 4:
                if strength_test_type == "Legs":
                    # Smaller screenshots fit on the same page - no page break, just a bit of space
                    doc.text.addElement(P(text=""))
                    doc.text.addElement(P(text=""))
                    doc.text.addElement(P(text="Kraftanalyse Vergleich Antagonist - Agonist", stylename="HeadingStyle"))
                else:
                    doc.text.addElement(P(text="Kraftanalyse Vergleich Antagonist - Agonist", stylename="HeadingPageBreakStyle"))
                doc.text.addElement(P(text=""))

                # Main screenshot (index 2)
                if kraft_screenshots[2] and os.path.exists(kraft_screenshots[2]):
                    try:
                        with PILImage.open(kraft_screenshots[2]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot3", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[2])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 3 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 3: {e}")

                # Flat bottom screenshot (index 3)
                if kraft_screenshots[3] and os.path.exists(kraft_screenshots[3]):
                    try:
                        with PILImage.open(kraft_screenshots[3]) as img:
                            img_width_px, img_height_px = img.size
                        frame_width_cm = 16.0
                        aspect_ratio = img_height_px / img_width_px if img_width_px > 0 else 1
                        frame_height_cm = frame_width_cm * aspect_ratio
                        centered_p = P(stylename="CenterParagraph")
                        frame = Frame(name="KraftScreenshot4", width=f"{frame_width_cm}cm", height=f"{frame_height_cm}cm", anchortype="paragraph")
                        href = doc.addPicture(kraft_screenshots[3])
                        if href:
                            frame.addElement(Image(href=href))
                            centered_p.addElement(frame)
                            doc.text.addElement(centered_p)
                            print(f"Kraft Screenshot 4 inserted: {frame_width_cm}cm x {frame_height_cm:.2f}cm")
                    except Exception as e:
                        print(f"Error adding kraft screenshot 4: {e}")

        # Add new section: Therapieempfehlungen (last page)
        doc.text.addElement(P(text="Therapieempfehlungen", stylename="HeadingPageBreakStyle"))
        doc.text.addElement(P(text=""))

        # Add bullet points from user input
        if therapie_texts and len(therapie_texts) > 0:
            therapie_bullet_list = List(stylename="BulletList")
            for text in therapie_texts:
                bullet_item = ListItem()
                bullet_item.addElement(P(text=text, stylename="BulletTextStyle"))
                therapie_bullet_list.addElement(bullet_item)
            doc.text.addElement(therapie_bullet_list)

        # Add spacing to move content to vertical middle of page
        for _ in range(13):
            doc.text.addElement(P(text=""))

        # New section: Umsetzung der Therapieempfehlung
        doc.text.addElement(P(text="Umsetzung der Therapieempfehlung", stylename="HeadingStyle"))
        doc.text.addElement(P(text=""))

        # First paragraph
        umsetzung_text = "Zur optimalen Umsetzung der Therapieempfehlung empfehlen wir, die weitere Behandlung direkt bei uns in der Praxis fortzuführen. Da wir den Befund im Bewegungslabor selbst erhoben haben, können wir die Therapie gezielt, strukturiert und ohne Informationsverlust planen und umsetzen."
        doc.text.addElement(P(text=umsetzung_text, stylename="BackgroundTextStyle"))

        # Bullet points for therapy options (indented)
        umsetzung_bullet_list = List(stylename="BulletList")

        # First bullet point: NMTT
        nmtt_text = "Als bevorzugte Option bieten wir eine neuromuskuläre Trainingstherapie (NMTT) im 1:1-Setting an, ein- bis zweimal pro Woche. Diese kombiniert – abhängig von Befund und Beschwerdebild – Huber-Training und die Egoscue-Methode, ein Trainingskonzept zur Korrektur von Fehlhaltungen. Durch die enge therapeutische Anleitung ist hier eine besonders effektive Korrektur von Dysbalancen möglich."
        nmtt_item = ListItem()
        nmtt_item.addElement(P(text=nmtt_text, stylename="IndentedBulletTextStyle"))
        umsetzung_bullet_list.addElement(nmtt_item)

        # Second bullet point: IMTT
        imtt_text = "Alternativ besteht die Möglichkeit einer individuellen medizinischen Trainingstherapie (IMTT) auf Basis der Egoscue-Methode für das eigenständige Training zu Hause. Die Übungen sind gerätefrei, alltagstauglich und werden im Rahmen monatlicher Termine regelmäßig angepasst. Wir bieten das IMTT auch als online Sitzung an, dafür eignet sich diese Option vor allem bei leichteren Beschwerden, eingeschränkter Zeit oder längerer Anfahrt; bei ausgeprägten Dysbalancen ist die NMTT klar zu bevorzugen."
        imtt_item = ListItem()
        imtt_item.addElement(P(text=imtt_text, stylename="IndentedBulletTextStyle"))
        umsetzung_bullet_list.addElement(imtt_item)

        doc.text.addElement(umsetzung_bullet_list)

    else:
        doc.text.addElement(P(text="Fehler beim Lesen der Messwerte aus der INI-Datei."))

    doc.save(odt_path)


class MeasurementTypeSelector:
    def __init__(self, parent):
        self.parent = parent
        self.measurement_type = None
        self.top = tk.Toplevel(parent)
        self.top.title("Measurement Protocol")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Measurement Protocol", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Select measurement type:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.type_var = tk.StringVar(value="Gehen")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("IOS", "IOS"), ("Statik", "Statik"), ("Gehen", "Gehen"), ("Laufen", "Laufen")]:
            tk.Radiobutton(options_frame, text=text, variable=self.type_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 320)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 320)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.measurement_type = self.type_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.measurement_type = None
        self.top.destroy()

    def get_measurement_type(self):
        self.parent.wait_window(self.top)
        return self.measurement_type


class GenderSelector:
    def __init__(self, parent):
        self.parent = parent
        self.gender = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Select Patient Gender")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Patient Gender", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Select gender:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.gender_var = tk.StringVar(value="Male")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Male", "Male"), ("Female", "Female")]:
            tk.Radiobutton(options_frame, text=text, variable=self.gender_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 280)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 280)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.gender = self.gender_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.gender = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_gender(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.gender


class TitleSelector:
    def __init__(self, parent):
        self.parent = parent
        self.academic_title = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Select Academic Title")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Academic Title", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Select title:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.title_var = tk.StringVar(value="None")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Prof Dr.", "Prof Dr."), ("Dr.", "Dr."), ("None", "None")]:
            tk.Radiobutton(options_frame, text=text, variable=self.title_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 300)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 300)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.academic_title = self.title_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.academic_title = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_title(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.academic_title


class ReportCreatorSelector:
    def __init__(self, parent):
        self.parent = parent
        self.creator_name = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Select Report Creator")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Report Creator", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Select the report creator:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.creator_var = tk.StringVar(value="Carlo Lade")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        names = ["Carlo Lade", "Linnea Nirenberg", "Clara Guntermann", "Valentin Stark", "Lena Ranzinger", "Johannes Paa"]
        for name in names:
            tk.Radiobutton(options_frame, text=name, variable=self.creator_var, value=name,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=3)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 350)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 350)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.creator_name = self.creator_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.creator_name = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_creator(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.creator_name


class StatikSIMSelector:
    def __init__(self, parent, default_selection="SIM"):
        self.parent = parent
        self.source = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("4D Average Source")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="4D Average Source", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Extract 4D average from static or SIM recording?", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.source_var = tk.StringVar(value=default_selection)

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Statik", "Statik"), ("SIM", "SIM")]:
            tk.Radiobutton(options_frame, text=text, variable=self.source_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(400, 280)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 400)
        window_height = max(self.top.winfo_height(), 280)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.source = self.source_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.source = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_source(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.source


class SIMMeasurementSelector:
    def __init__(self, parent):
        self.parent = parent
        self.sim_performed = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("SIM Measurement")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="SIM Measurement", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Was SIM measurement performed?", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.sim_var = tk.StringVar(value="Nein")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Yes", "Ja"), ("No", "Nein")]:
            tk.Radiobutton(options_frame, text=text, variable=self.sim_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(380, 280)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 380)
        window_height = max(self.top.winfo_height(), 280)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.sim_performed = self.sim_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.sim_performed = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_sim_status(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.sim_performed


class LegLengthSelector:
    """Dialog for selecting whether to include leg length examination results"""
    def __init__(self, parent):
        self.parent = parent
        self.leg_length_selected = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Leg Length Examination")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Leg Length Examination", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Export leg length examination test results?", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.leg_length_var = tk.StringVar(value="Nein")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Yes", "Ja"), ("No", "Nein")]:
            tk.Radiobutton(options_frame, text=text, variable=self.leg_length_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(380, 280)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 380)
        window_height = max(self.top.winfo_height(), 280)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.leg_length_selected = self.leg_length_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.leg_length_selected = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_leg_length_status(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.leg_length_selected


class StrengthTestTypeSelector:
    """Dialog for selecting the type of strength test performed"""
    def __init__(self, parent):
        self.parent = parent
        self.strength_test_type = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Strength Test Type")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Strength Test Type", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Select strength test type:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.type_var = tk.StringVar(value="Torso + legs")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text in ["Torso + legs", "Torso + legs + shoulders", "Torso + shoulders", "Legs + shoulders", "Legs"]:
            tk.Radiobutton(options_frame, text=text, variable=self.type_var, value=text,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(380, 320)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 380)
        window_height = max(self.top.winfo_height(), 320)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.strength_test_type = self.type_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.strength_test_type = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_strength_test_type(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.strength_test_type


class ISGSelector:
    def __init__(self, parent):
        self.parent = parent
        self.isg_right = None
        self.isg_left = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("ISG Status")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="ISG Status", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)

        # Right ISG
        tk.Label(main_frame, text="ISG right:", font=("Helvetica", 11, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(5, 0))
        self.right_var = tk.StringVar(value="Frei")
        right_frame = tk.Frame(main_frame, bg="#F5F5F5")
        right_frame.pack(pady=5)
        for text, value in [("Free", "Frei"), ("Blocked", "blockiert")]:
            tk.Radiobutton(right_frame, text=text, variable=self.right_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20)

        # Left ISG
        tk.Label(main_frame, text="ISG left:", font=("Helvetica", 11, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(10, 0))
        self.left_var = tk.StringVar(value="Frei")
        left_frame = tk.Frame(main_frame, bg="#F5F5F5")
        left_frame.pack(pady=5)
        for text, value in [("Free", "Frei"), ("Blocked", "blockiert")]:
            tk.Radiobutton(left_frame, text=text, variable=self.left_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 350)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 350)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.isg_right = self.right_var.get()
        self.isg_left = self.left_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.isg_right = None
        self.isg_left = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_isg_status(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK", "BACK"
        return self.isg_right, self.isg_left


class MarkerSelector:
    def __init__(self, parent):
        self.parent = parent
        self.markers = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Marker Placement")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Marker Placement", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Markers placed?", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.dl_dr_var = tk.BooleanVar(value=False)
        self.ws_var = tk.BooleanVar(value=False)
        self.vp_var = tk.BooleanVar(value=False)
        self.keine_var = tk.BooleanVar(value=False)

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, var in [("DL & DR", self.dl_dr_var), ("WS", self.ws_var), ("VP", self.vp_var), ("None", self.keine_var)]:
            tk.Checkbutton(options_frame, text=text, variable=var,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=3)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(350, 320)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 350)
        window_height = max(self.top.winfo_height(), 320)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

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

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_markers(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.markers


class BeckenhochstandSelector:
    def __init__(self, parent):
        self.parent = parent
        self.beckenhochstand_side = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Pelvic Elevation Position")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Pelvic Elevation Position", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="For the Pelvic Drop calculation,\nis the pelvic elevation right, left, or level?",
                 font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.side_var = tk.StringVar(value="gerade")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Right", "rechts"), ("Left", "links"), ("Level", "gerade")]:
            tk.Radiobutton(options_frame, text=text, variable=self.side_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(400, 320)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 400)
        window_height = max(self.top.winfo_height(), 320)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.beckenhochstand_side = self.side_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.beckenhochstand_side = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_beckenhochstand_side(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.beckenhochstand_side


class CustomPelvicDropSelector:
    """Dialog to ask if user wants to enter custom pelvic drop values"""
    def __init__(self, parent):
        self.parent = parent
        self.use_custom = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Pelvic Drop Values")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Pelvic Drop Values", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Enter custom Pelvic Drop values\nfor uneven Pelvic Drop graph?",
                 font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.custom_var = tk.StringVar(value="Nein")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("Yes", "Ja"), ("No", "Nein")]:
            tk.Radiobutton(options_frame, text=text, variable=self.custom_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(380, 300)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 380)
        window_height = max(self.top.winfo_height(), 300)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.use_custom = self.custom_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.use_custom = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_choice(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.use_custom


class CustomPelvicDropInputDialog:
    """Dialog for entering custom L and R pelvic drop values"""
    def __init__(self, parent):
        self.parent = parent
        self.values = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Enter Pelvic Drop Values")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Pelvic Drop Values", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Please enter the Pelvic Drop values:", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        # L value input
        l_frame = tk.Frame(main_frame, bg="#F5F5F5")
        l_frame.pack(fill="x", padx=20, pady=8)
        tk.Label(l_frame, text="L (Left):", width=15, anchor="w", font=("Helvetica", 11), bg="#F5F5F5", fg="#333333").pack(side="left")
        self.l_entry = tk.Entry(l_frame, width=20, font=("Helvetica", 11), relief=tk.SOLID, bd=1)
        self.l_entry.pack(side="left", padx=5)

        # R value input
        r_frame = tk.Frame(main_frame, bg="#F5F5F5")
        r_frame.pack(fill="x", padx=20, pady=8)
        tk.Label(r_frame, text="R (Right):", width=15, anchor="w", font=("Helvetica", 11), bg="#F5F5F5", fg="#333333").pack(side="left")
        self.r_entry = tk.Entry(r_frame, width=20, font=("Helvetica", 11), relief=tk.SOLID, bd=1)
        self.r_entry.pack(side="left", padx=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(400, 280)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 400)
        window_height = max(self.top.winfo_height(), 280)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        try:
            l_value = float(self.l_entry.get().strip().replace(',', '.'))
            r_value = float(self.r_entry.get().strip().replace(',', '.'))
            self.values = (l_value, r_value)
            self.top.destroy()
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values.")

    def _on_cancel(self):
        self.values = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_values(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.values


class BulletPointInputDialog:
    """Dialog for entering multiple bullet point texts"""
    def __init__(self, parent, title, num_fields=5, show_numbers=True):
        self.parent = parent
        self.texts = None
        self.went_back = False
        self.num_fields = num_fields
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text=title, font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="(Empty fields will be skipped)", font=("Helvetica", 9), fg="#afa190", bg="#F5F5F5").pack(pady=(0, 15))

        self.entries = []
        for i in range(num_fields):
            frame = tk.Frame(main_frame, bg="#F5F5F5")
            frame.pack(fill="x", padx=10, pady=5)
            if show_numbers:
                tk.Label(frame, text=f"{i+1}.", width=3, font=("Helvetica", 11), bg="#F5F5F5", fg="#333333").pack(side="left")
            entry = tk.Entry(frame, width=55, font=("Helvetica", 11), relief=tk.SOLID, bd=1)
            entry.pack(side="left", fill="x", expand=True)
            self.entries.append(entry)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        # Calculate min height based on number of fields (base height + per field)
        min_height = 200 + (self.num_fields * 40)
        self.top.minsize(550, min_height)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 550)
        window_height = max(self.top.winfo_height(), min_height)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        # Only collect non-empty entries
        self.texts = [entry.get().strip() for entry in self.entries if entry.get().strip()]
        self.top.destroy()

    def _on_cancel(self):
        self.texts = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_texts(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.texts


class ExportFormatSelector:
    """Dialog for selecting export format (PDF, ODT, or both)"""
    def __init__(self, parent):
        self.parent = parent
        self.format = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title("Select Export Format")
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text="Export Format", font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text="Which file format should be created?", font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.format_var = tk.StringVar(value="ODT")

        options_frame = tk.Frame(main_frame, bg="#F5F5F5")
        options_frame.pack(pady=10)

        for text, value in [("PDF", "PDF"), ("ODT", "ODT"), ("Both (PDF and ODT)", "BOTH")]:
            tk.Radiobutton(options_frame, text=text, variable=self.format_var, value=value,
                          font=("Helvetica", 11), bg="#F5F5F5", fg="#333333",
                          activebackground="#F5F5F5", selectcolor="#FFFFFF").pack(anchor="w", padx=20, pady=5)

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(380, 300)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 380)
        window_height = max(self.top.winfo_height(), 300)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.format = self.format_var.get()
        self.top.destroy()

    def _on_cancel(self):
        self.format = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_format(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.format


class TextInputDialog:
    """Dialog for entering text with Back button support"""
    def __init__(self, parent, title, prompt, initial_value=""):
        self.parent = parent
        self.result = None
        self.went_back = False
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg="#F5F5F5")
        self.top.transient(parent)
        self.top.grab_set()

        main_frame = tk.Frame(self.top, bg="#F5F5F5")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        tk.Label(main_frame, text=title, font=("Helvetica", 14, "bold"), fg="#333333", bg="#F5F5F5").pack(pady=(0, 5))
        # Teal-brown separator line
        sep_frame = tk.Frame(main_frame, bg="#F5F5F5", height=3)
        sep_frame.pack(fill=tk.X, pady=(5, 15))
        tk.Frame(sep_frame, bg="#80afaa", height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Frame(sep_frame, bg="#afa190", height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        tk.Label(main_frame, text=prompt, font=("Helvetica", 11), fg="#333333", bg="#F5F5F5").pack(pady=(0, 15))

        self.entry = tk.Entry(main_frame, width=45, font=("Helvetica", 11), relief=tk.SOLID, bd=1)
        self.entry.pack(pady=10, padx=20)
        if initial_value:
            self.entry.insert(0, initial_value)
        self.entry.focus_set()

        button_frame = tk.Frame(main_frame, bg="#F5F5F5")
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="Back", command=self._on_back, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="OK", command=self._on_ok, width=10, font=("Helvetica", 10),
                 bg="#4CAF50", fg="#FFFFFF", activebackground="#43A047", activeforeground="#FFFFFF",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)
        tk.Button(button_frame, text="Cancel", command=self._on_cancel, width=10, font=("Helvetica", 10),
                 bg="#E57373", fg="#333333", activebackground="#EF5350", activeforeground="#333333",
                 relief=tk.FLAT, cursor="hand2").pack(side="left", padx=5)

        self.entry.bind("<Return>", lambda e: self._on_ok())
        self._center_window()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _center_window(self):
        self.top.update_idletasks()
        self.top.minsize(450, 220)
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        window_width = max(self.top.winfo_width(), 450)
        window_height = max(self.top.winfo_height(), 220)
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.top.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def _on_ok(self):
        self.result = self.entry.get().strip()
        self.top.destroy()

    def _on_cancel(self):
        self.result = None
        self.top.destroy()

    def _on_back(self):
        self.went_back = True
        self.top.destroy()

    def get_input(self):
        self.parent.wait_window(self.top)
        if self.went_back:
            return "BACK"
        return self.result


class CoordinateFinder:
    """Base class for finding coordinates from PDF screenshots"""
    def __init__(self, parent):
        self.parent = parent
        self.image = None
        self.photo = None
        self.clicks = []
        self.canvas = None
        self.top = None
        self.scale = 1.0
        self.final_coordinates = None

    def find_coordinates_from_path(self, pdf_path, page_num=1, screenshot_name="Screenshot"):
        """Find coordinates from a specific page of a PDF"""
        try:
            print(f"Converting PDF page {page_num} to image...")
            images = convert_from_path(pdf_path, dpi=150, first_page=page_num, last_page=page_num)
            if not images:
                messagebox.showerror("Error", f"Could not convert PDF page {page_num} to image")
                return

            self.image = images[0]
            self.clicks = []
            print(f"Image size: {self.image.size[0]}x{self.image.size[1]} pixels (at 150 DPI)")

            self.top = tk.Toplevel(self.parent)
            window_title = f"{screenshot_name} - Page {page_num} - Click Top-Left, then Bottom-Right"
            self.top.title(window_title)
            self.top.configure(bg="#F5F5F5")
            self.top.transient(self.parent)
            self.top.grab_set()

            max_width = 1200
            max_height = 800
            img_width, img_height = self.image.size

            self.scale = 1.0
            if img_width > max_width or img_height > max_height:
                scale_w = max_width / img_width
                scale_h = max_height / img_height
                self.scale = min(scale_w, scale_h)
                display_width = int(img_width * self.scale)
                display_height = int(img_height * self.scale)
                display_image = self.image.resize((display_width, display_height), PILImage.Resampling.LANCZOS)
            else:
                display_image = self.image
                display_width = img_width
                display_height = img_height

            # Header label
            tk.Label(self.top, text="Click 1: Top-Left corner, Click 2: Bottom-Right corner",
                    font=("Helvetica", 12, "bold"), fg="#80afaa", bg="#F5F5F5").pack(pady=10)

            self.canvas = tk.Canvas(self.top, width=display_width, height=display_height, bg="#FFFFFF")
            self.canvas.pack(padx=10, pady=(0, 10))

            self.photo = ImageTk.PhotoImage(display_image)
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

            self.canvas.bind("<Button-1>", self.on_click)

        except Exception as e:
            messagebox.showerror("Error", f"Error processing PDF: {e}")
            import traceback
            traceback.print_exc()

    def wait_for_completion(self):
        if self.top:
            self.parent.wait_window(self.top)

    def on_click(self, event):
        x_display = event.x
        y_display = event.y
        x_original = int(x_display / self.scale)
        y_original = int(y_display / self.scale)

        self.clicks.append((x_original, y_original))

        r = 5
        self.canvas.create_oval(x_display-r, y_display-r, x_display+r, y_display+r,
                                fill="red", outline="white", width=2)
        self.canvas.create_text(x_display, y_display-15, text=f"Click {len(self.clicks)}",
                                fill="red", font=("Arial", 10, "bold"))

        print(f"Click {len(self.clicks)}: Display({x_display}, {y_display}) -> Original({x_original}, {y_original})")

        if len(self.clicks) == 2:
            self.calculate_coordinates()

    def calculate_coordinates(self):
        x1, y1 = self.clicks[0]
        x2, y2 = self.clicks[1]

        left = min(x1, x2)
        top = min(y1, y2)
        right = max(x1, x2)
        bottom = max(y1, y2)

        img_width, img_height = self.image.size

        left_pct = (left / img_width) * 100
        top_pct = (top / img_height) * 100
        right_pct = (right / img_width) * 100
        bottom_pct = (bottom / img_height) * 100

        x1_display = int(left * self.scale)
        y1_display = int(top * self.scale)
        x2_display = int(right * self.scale)
        y2_display = int(bottom * self.scale)
        self.canvas.create_rectangle(x1_display, y1_display, x2_display, y2_display,
                                      outline="green", width=3)

        self.final_coordinates = {
            'left_pct': left_pct,
            'top_pct': top_pct,
            'right_pct': right_pct,
            'bottom_pct': bottom_pct
        }

        result = f"""
CROP COORDINATES FOUND:
Left: {left_pct:.2f}%
Top: {top_pct:.2f}%
Right: {right_pct:.2f}%
Bottom: {bottom_pct:.2f}%
"""
        print(result)
        messagebox.showinfo("Coordinates Found", result)
        self.top.destroy()


class RunningCoordinateFinder:
    """Coordinate finder for running analysis screenshots from hp.pdf and ios.pdf"""
    def __init__(self, parent):
        self.parent = parent
        self.screenshots_config = [
            {"name": "Running Screenshot 1", "page": 1, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 2", "page": 2, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 3", "page": 3, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 4", "page": 6, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 5", "page": 4, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 6", "page": 5, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 7", "page": 8, "pdf": "hp.pdf"},
            {"name": "Running Screenshot 8", "page": 1, "pdf": "ios.pdf"},
        ]
        self.current_screenshot_index = 0
        self.all_coordinates = []

    def find_all_coordinates(self):
        folder_path = filedialog.askdirectory(title="Select folder containing hp.pdf and ios.pdf")
        if not folder_path:
            messagebox.showinfo("Cancelled", "Folder selection was cancelled.")
            return

        hp_path = os.path.join(folder_path, "hp.pdf")
        ios_path = os.path.join(folder_path, "ios.pdf")

        if not os.path.exists(hp_path):
            messagebox.showerror("Error", f"Could not find hp.pdf in {folder_path}")
            return
        if not os.path.exists(ios_path):
            messagebox.showerror("Error", f"Could not find ios.pdf in {folder_path}")
            return

        self.folder_path = folder_path
        self.process_next_screenshot()

    def process_next_screenshot(self):
        if self.current_screenshot_index >= len(self.screenshots_config):
            self.show_all_results()
            return

        config = self.screenshots_config[self.current_screenshot_index]
        pdf_path = os.path.join(self.folder_path, config['pdf'])

        print(f"\n{'='*60}")
        print(f"Processing {config['name']}: Page {config['page']} from {config['pdf']}")
        print(f"{'='*60}")

        finder = CoordinateFinder(self.parent)
        finder.find_coordinates_from_path(pdf_path, config['page'], config['name'])
        finder.wait_for_completion()

        if finder.final_coordinates:
            self.all_coordinates.append({
                'screenshot': config['name'],
                'page': config['page'],
                'pdf': config['pdf'],
                'coordinates': finder.final_coordinates
            })

        self.current_screenshot_index += 1
        self.process_next_screenshot()

    def show_all_results(self):
        print("\n" + "="*70)
        print("RUNNING ANALYSIS COORDINATES (hp.pdf & ios.pdf)")
        print("="*70)

        result_text = "RUNNING ANALYSIS COORDINATES:\n" + "="*50 + "\n\n"

        for coord_data in self.all_coordinates:
            coords = coord_data['coordinates']
            text = f"{coord_data['screenshot']} (Page {coord_data['page']} from {coord_data['pdf']}):\n"
            text += f"  Left: {coords['left_pct']:.2f}%\n"
            text += f"  Top: {coords['top_pct']:.2f}%\n"
            text += f"  Right: {coords['right_pct']:.2f}%\n"
            text += f"  Bottom: {coords['bottom_pct']:.2f}%\n\n"

            print(text)
            result_text += text

        messagebox.showinfo("All Coordinates Found", result_text)


def find_running_coordinates():
    """Helper function to find all running analysis coordinates"""
    finder = RunningCoordinateFinder(root)
    finder.find_all_coordinates()


def find_coordinates():
    """Helper function to find coordinates from any PDF page"""
    # Ask user to select a PDF file
    pdf_path = filedialog.askopenfilename(
        title="Select PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return

    # Ask for page number
    page_num = simpledialog.askinteger(
        "Page Number",
        "Enter the page number to extract coordinates from:",
        minvalue=1,
        initialvalue=1
    )
    if page_num is None:
        return

    # Open coordinate finder
    finder = CoordinateFinder(root)
    finder.find_coordinates_from_path(pdf_path, page_num, f"Page {page_num}")


def generate_report():
    # Data dictionary to store all collected values
    data = {}

    # Step index for navigation
    current_step = 0

    # Define the steps (some are conditional)
    # Steps: 0=measurement_type, 1=report_creator, 2=statik_sim_source, 3=gender, 4=title,
    #        5=markers, 6=sim_performed, 7=isg (conditional), 8=leg_length,
    #        9=strength_test_type, 10=folder_path (extracts patient info),
    #        11=beckenhochstand (conditional), 12=leg_length_texts, 13=beinachsen,
    #        14=ganganalyse (conditional), 15=therapie, 16=export_format, 17=save_path

    while True:
        if current_step == 0:
            # Measurement type selection
            measurement_type_selector = MeasurementTypeSelector(root)
            result = measurement_type_selector.get_measurement_type()

            if result is None:
                messagebox.showinfo("Cancelled", "Measurement type selection was cancelled.")
                return
            # Can't go back from step 0
            data['measurement_type'] = result
            print(f"Selected measurement type: {result}")
            current_step += 1

        elif current_step == 1:
            # Report creator selection
            creator_selector = ReportCreatorSelector(root)
            result = creator_selector.get_creator()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Report creator selection was cancelled.")
                return
            data['report_creator'] = result
            current_step += 1

        elif current_step == 2:
            # Auto-detect which PDF to use (4d_average.pdf or statik.pdf)
            # This will be used later in step 8
            current_step += 1

        elif current_step == 3:
            # Gender selection
            gender_selector = GenderSelector(root)
            result = gender_selector.get_gender()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Gender selection was cancelled.")
                return
            data['gender'] = result
            current_step += 1

        elif current_step == 4:
            # Academic title selection
            title_selector = TitleSelector(root)
            result = title_selector.get_title()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Academic title selection was cancelled.")
                return
            data['academic_title'] = result
            current_step += 1

        elif current_step == 5:
            # Marker selection
            marker_selector = MarkerSelector(root)
            result = marker_selector.get_markers()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Marker selection was cancelled.")
                return
            data['markers'] = result
            current_step += 1

        elif current_step == 6:
            # SIM measurement selection
            sim_selector = SIMMeasurementSelector(root)
            result = sim_selector.get_sim_status()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "SIM measurement selection was cancelled.")
                return
            data['sim_performed'] = result
            current_step += 1

        elif current_step == 7:
            # ISG status (conditional - only if SIM was performed)
            if data.get('sim_performed') == "Ja":
                isg_selector = ISGSelector(root)
                isg_right, isg_left = isg_selector.get_isg_status()

                if isg_right == "BACK":
                    current_step -= 1
                    continue
                if isg_right is None or isg_left is None:
                    messagebox.showinfo("Cancelled", "ISG status selection was cancelled.")
                    return
                data['isg_right'] = isg_right
                data['isg_left'] = isg_left
            else:
                data['isg_right'] = None
                data['isg_left'] = None
            current_step += 1

        elif current_step == 8:
            # Leg length examination selection
            leg_length_selector = LegLengthSelector(root)
            result = leg_length_selector.get_leg_length_status()

            if result == "BACK":
                # If SIM was not performed, go back to SIM step, otherwise go back to ISG step
                if data.get('sim_performed') != "Ja":
                    current_step -= 2
                else:
                    current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Leg length examination selection was cancelled.")
                return
            data['leg_length_selected'] = result
            current_step += 1

        elif current_step == 9:
            # Strength test type selection (only for Statik, Gehen, Laufen)
            measurement_type = data['measurement_type']
            if measurement_type != "IOS":
                strength_selector = StrengthTestTypeSelector(root)
                result = strength_selector.get_strength_test_type()

                if result == "BACK":
                    current_step -= 1
                    continue
                if result is None:
                    messagebox.showinfo("Cancelled", "Strength test type selection was cancelled.")
                    return
                data['strength_test_type'] = result
            else:
                data['strength_test_type'] = "Torso + legs"
            current_step += 1

        elif current_step == 10:
            # Folder selection - this step has no Back button in native dialog
            folder_path = filedialog.askdirectory(title="Select folder containing measurement files")
            if not folder_path:
                # User cancelled - ask if they want to go back
                if messagebox.askyesno("Go Back?", "Folder selection was cancelled. Go back to previous step?"):
                    # Skip strength test type step for IOS (go back to leg length)
                    if data['measurement_type'] == "IOS":
                        current_step -= 2
                    else:
                        current_step -= 1
                    continue
                else:
                    return
            data['folder_path'] = folder_path

            # Process files from the folder (automatic, no user interaction)
            # Auto-find .ini file containing "4D average" (case-insensitive)
            ini_path = None
            for filename in os.listdir(folder_path):
                if filename.lower().endswith('.ini') and '4d average' in filename.lower():
                    ini_path = os.path.join(folder_path, filename)
                    break

            if not ini_path:
                messagebox.showerror("Error", "Could not find .ini file containing '4D average' in the selected folder.")
                # Let user select a different folder
                continue

            data['ini_path'] = ini_path
            print(f"Found .ini file: {ini_path}")

            # Auto-detect which PDF to use: prioritize 4d_average.pdf, fallback to statik.pdf
            pdf_path_4d = os.path.join(folder_path, "4d_average.pdf")
            pdf_path_statik = os.path.join(folder_path, "statik.pdf")

            if os.path.exists(pdf_path_4d):
                pdf_path = pdf_path_4d
                print(f"Using 4d_average.pdf for 4D Average screenshot: {pdf_path}")
            elif os.path.exists(pdf_path_statik):
                pdf_path = pdf_path_statik
                print(f"4d_average.pdf not found, using statik.pdf for 4D Average screenshot: {pdf_path}")
            else:
                messagebox.showerror("Error", f"Could not find 4d_average.pdf or statik.pdf in {folder_path}")
                continue

            # Extract patient info from the selected PDF
            patient_name, patient_dob, measurement_date = extract_patient_info_from_pdf(pdf_path)
            if patient_name:
                data['patient_name'] = patient_name
                print(f"Extracted patient name: {patient_name}")
            else:
                messagebox.showerror("Error", "Could not extract patient name from PDF")
                continue
            if patient_dob:
                data['patient_dob'] = patient_dob
                print(f"Extracted patient DOB: {patient_dob}")
            else:
                messagebox.showerror("Error", "Could not extract patient date of birth from PDF")
                continue
            if measurement_date:
                data['measurement_date'] = measurement_date
                print(f"Extracted measurement date: {measurement_date}")
            else:
                messagebox.showerror("Error", "Could not extract measurement date from PDF")
                continue

            # Crop screenshot from the selected PDF (same coordinates, page 1)
            screenshot_path = crop_pdf_screenshot(pdf_path)
            if not screenshot_path:
                messagebox.showerror("Error", "Failed to crop screenshot from PDF")
                continue

            data['screenshot_path'] = screenshot_path

            # Auto-find and crop screenshots from statik.pdf
            measurement_type = data['measurement_type']
            statik_pdf_path = os.path.join(folder_path, "statik.pdf")
            statik_screenshots = []
            if os.path.exists(statik_pdf_path):
                print(f"Found statik.pdf: {statik_pdf_path}")
                statik_screenshots = crop_statik_screenshots(statik_pdf_path)
                if not statik_screenshots or len(statik_screenshots) < 7:
                    print("Warning: Failed to crop all statik screenshots")
            else:
                print("Warning: statik.pdf not found in folder")
            data['statik_screenshots'] = statik_screenshots

            # Auto-find and crop screenshots based on measurement type
            gehen_screenshots = []
            ios_pedografie_screenshots = []
            if measurement_type == "Gehen":
                gehen_pdf_path = os.path.join(folder_path, "gehen.pdf")
                if os.path.exists(gehen_pdf_path):
                    print(f"Found gehen.pdf: {gehen_pdf_path}")
                    gehen_screenshots = crop_gehen_screenshots(gehen_pdf_path)
                    if not gehen_screenshots or len(gehen_screenshots) < 8:
                        print("Warning: Failed to crop all gehen screenshots")
                else:
                    print("Warning: gehen.pdf not found in folder")
            elif measurement_type == "Laufen":
                hp_pdf_path = os.path.join(folder_path, "hp.pdf")
                ios_pdf_path = os.path.join(folder_path, "ios.pdf")
                if os.path.exists(hp_pdf_path) and os.path.exists(ios_pdf_path):
                    print(f"Found hp.pdf: {hp_pdf_path}")
                    print(f"Found ios.pdf: {ios_pdf_path}")
                    gehen_screenshots = crop_laufen_screenshots(hp_pdf_path, ios_pdf_path)
                    if not gehen_screenshots or len(gehen_screenshots) < 8:
                        print("Warning: Failed to crop all laufen screenshots")
                else:
                    if not os.path.exists(hp_pdf_path):
                        print("Warning: hp.pdf not found in folder")
                    if not os.path.exists(ios_pdf_path):
                        print("Warning: ios.pdf not found in folder")
            elif measurement_type in ["Statik", "IOS"]:
                ios_pdf_path = os.path.join(folder_path, "ios.pdf")
                if os.path.exists(ios_pdf_path):
                    print(f"Found ios.pdf: {ios_pdf_path}")
                    ios_pedografie_screenshots = crop_ios_pedografie_screenshots(ios_pdf_path)
                    if not ios_pedografie_screenshots or len(ios_pedografie_screenshots) < 2:
                        print("Warning: Failed to crop all ios pedografie screenshots")
                else:
                    print("Warning: ios.pdf not found in folder")

            data['gehen_screenshots'] = gehen_screenshots
            data['ios_pedografie_screenshots'] = ios_pedografie_screenshots

            # Auto-find and crop screenshots from kraft.pdf (not for IOS type)
            kraft_pdf_path = os.path.join(folder_path, "kraft.pdf")
            kraft_screenshots = []
            if measurement_type != "IOS":
                if os.path.exists(kraft_pdf_path):
                    print(f"Found kraft.pdf: {kraft_pdf_path}")
                    strength_test_type = data.get('strength_test_type', 'Torso + legs')
                    kraft_screenshots = crop_kraft_screenshots(kraft_pdf_path, strength_test_type)
                    expected_count = 6 if strength_test_type == "Torso + legs + shoulders" else 4
                    if not kraft_screenshots or len(kraft_screenshots) < expected_count:
                        print("Warning: Failed to crop all kraft screenshots")
                else:
                    print("Warning: kraft.pdf not found in folder")
            data['kraft_screenshots'] = kraft_screenshots

            # Process vgl.pdf for leg length examination (conditional)
            vgl_screenshot = None
            if data.get('leg_length_selected') == "Ja":
                vgl_pdf_path = os.path.join(folder_path, "vgl.pdf")
                if os.path.exists(vgl_pdf_path):
                    print(f"Found vgl.pdf: {vgl_pdf_path}")
                    vgl_screenshot = crop_vgl_screenshot(vgl_pdf_path)
                    if not vgl_screenshot:
                        print("Warning: Failed to crop vgl.pdf screenshot")
                else:
                    messagebox.showwarning("Missing File",
                        "vgl.pdf not found in folder.\n\n"
                        "Please add the vgl.pdf file to the folder and select the folder again.")
                    # Go back to folder selection
                    continue
            data['vgl_screenshot'] = vgl_screenshot

            current_step += 1

        elif current_step == 11:
            # Beckenhochstand (conditional - only for Gehen/Laufen with motion data)
            measurement_type = data['measurement_type']
            folder_path = data['folder_path']
            ini_path = data['ini_path']
            data['pelvic_drop_sentence'] = None

            if measurement_type in ["Gehen", "Laufen"]:
                motion_ini_path = None
                for filename in os.listdir(folder_path):
                    if filename.lower().endswith('.ini') and '4d motion' in filename.lower():
                        motion_ini_path = os.path.join(folder_path, filename)
                        break

                if motion_ini_path:
                    print(f"Found motion .ini file: {motion_ini_path}")
                    (kyphosis_angle, lordosis_angle, scoliosis_angle,
                     surf_rot_left, surf_rot_right, lat_dev_left, lat_dev_right, sva_axis, beckenhochstand) = parse_ini_file(ini_path)
                    motion_mean, motion_min, motion_max = parse_motion_ini_file(motion_ini_path)

                    if beckenhochstand is not None and motion_mean is not None and motion_min is not None and motion_max is not None:
                        data['_beckenhochstand'] = beckenhochstand
                        data['_motion_mean'] = motion_mean
                        data['_motion_min'] = motion_min
                        data['_motion_max'] = motion_max

                        # Ask if user wants to enter custom values
                        custom_selector = CustomPelvicDropSelector(root)
                        custom_choice = custom_selector.get_choice()

                        if custom_choice == "BACK":
                            current_step -= 1
                            continue
                        if custom_choice is None:
                            messagebox.showinfo("Cancelled", "Custom pelvic drop selection was cancelled.")
                            return

                        if custom_choice == "Ja":
                            # Get custom L and R values
                            custom_input = CustomPelvicDropInputDialog(root)
                            custom_values = custom_input.get_values()

                            if custom_values == "BACK":
                                current_step -= 1
                                continue
                            if custom_values is None:
                                messagebox.showinfo("Cancelled", "Custom pelvic drop input was cancelled.")
                                return

                            l_value, r_value = custom_values
                            pelvic_drop_sentence = generate_pelvic_drop_sentence_from_custom(r_value, l_value)
                            print(f"Using custom pelvic drop values: R={r_value}, L={l_value}")
                        else:
                            # Use automatic calculation
                            pelvic_drop_sentence = calculate_pelvic_drop_sentence(
                                beckenhochstand, motion_mean, motion_min, motion_max
                            )

                        data['pelvic_drop_sentence'] = pelvic_drop_sentence
                        print(f"Pelvic drop sentence: {pelvic_drop_sentence}")
                    else:
                        print("Warning: Could not extract all values for pelvic drop calculation")
                else:
                    print("Warning: 4D motion .ini file not found in folder")
            current_step += 1

        elif current_step == 12:
            # Leg length examination textboxes (conditional)
            if data.get('leg_length_selected') == "Ja" and data.get('vgl_screenshot'):
                leg_length_dialog = BulletPointInputDialog(root, "Leg length examination findings", num_fields=2)
                result = leg_length_dialog.get_texts()

                if result == "BACK":
                    current_step -= 1
                    continue
                if result is None:
                    messagebox.showinfo("Cancelled", "Leg length examination findings was cancelled.")
                    return
                data['leg_length_texts'] = result
            else:
                data['leg_length_texts'] = None
            current_step += 1

        elif current_step == 13:
            # Beinachsen description
            beinachsen_dialog = BulletPointInputDialog(root, "Leg axis and posture analysis description", num_fields=3)
            result = beinachsen_dialog.get_texts()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Leg axis description was cancelled.")
                return
            data['beinachsen_texts'] = result
            current_step += 1

        elif current_step == 14:
            # Ganganalyse description (conditional - only for Gehen/Laufen)
            measurement_type = data['measurement_type']
            if measurement_type in ["Gehen", "Laufen"]:
                ganganalyse_title = "Running analysis description" if measurement_type == "Laufen" else "Gait analysis description"
                ganganalyse_dialog = BulletPointInputDialog(root, ganganalyse_title, num_fields=5)
                result = ganganalyse_dialog.get_texts()

                if result == "BACK":
                    current_step -= 1
                    continue
                if result is None:
                    abort_message = "Running analysis description was cancelled." if measurement_type == "Laufen" else "Gait analysis description was cancelled."
                    messagebox.showinfo("Cancelled", abort_message)
                    return
                data['ganganalyse_texts'] = result
            else:
                data['ganganalyse_texts'] = None
            current_step += 1

        elif current_step == 15:
            # Therapie recommendations
            therapie_dialog = BulletPointInputDialog(root, "Therapy recommendations", num_fields=5)
            result = therapie_dialog.get_texts()

            if result == "BACK":
                # If ganganalyse was skipped, go back to beinachsen
                if data['measurement_type'] not in ["Gehen", "Laufen"]:
                    current_step -= 2
                else:
                    current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Therapy recommendations was cancelled.")
                return
            data['therapie_texts'] = result
            current_step += 1

        elif current_step == 16:
            # Export format selection
            export_selector = ExportFormatSelector(root)
            result = export_selector.get_format()

            if result == "BACK":
                current_step -= 1
                continue
            if result is None:
                messagebox.showinfo("Cancelled", "Export format selection was cancelled.")
                return
            data['export_format'] = result
            current_step += 1

        elif current_step == 17:
            # Auto-generate save path using folder_path and patient last name
            export_format = data['export_format']
            folder_path = data['folder_path']
            patient_name = data['patient_name']

            # Extract last name (patient_name is in "Nachname, Vorname" format)
            if ',' in patient_name:
                last_name = patient_name.split(',')[0].strip()
            else:
                last_name = patient_name.strip()

            # Generate filename: "<YYYY-MM-DD> Motionlab Report <last name>"
            current_date = datetime.now()
            date_iso = current_date.strftime("%Y-%m-%d")
            base_filename = f"{date_iso} Motionlab Report {last_name}"

            if export_format == "ODT":
                save_path = os.path.join(folder_path, f"{base_filename}.odt")
            else:
                save_path = os.path.join(folder_path, f"{base_filename}.pdf")

            data['save_path'] = save_path
            # All steps completed, break the loop
            break

    # Extract all collected data
    measurement_type = data['measurement_type']
    measurement_date = data['measurement_date']
    gender = data['gender']
    academic_title = data['academic_title']
    patient_name = data['patient_name']
    patient_dob = data['patient_dob']
    report_creator = data['report_creator']
    folder_path = data['folder_path']
    ini_path = data['ini_path']
    screenshot_path = data['screenshot_path']
    statik_screenshots = data['statik_screenshots']
    gehen_screenshots = data['gehen_screenshots']
    ios_pedografie_screenshots = data['ios_pedografie_screenshots']
    kraft_screenshots = data['kraft_screenshots']
    strength_test_type = data.get('strength_test_type', 'Torso + legs')
    sim_performed = data['sim_performed']
    isg_right = data['isg_right']
    isg_left = data['isg_left']
    markers = data['markers']
    pelvic_drop_sentence = data['pelvic_drop_sentence']
    beinachsen_texts = data['beinachsen_texts']
    ganganalyse_texts = data['ganganalyse_texts']
    therapie_texts = data['therapie_texts']
    leg_length_texts = data.get('leg_length_texts')
    vgl_screenshot = data.get('vgl_screenshot')
    export_format = data['export_format']
    save_path = data['save_path']

    # Build patient full title
    salutation = "Herr" if gender == "Male" else "Frau"
    patient_full_title = salutation
    if academic_title and academic_title != "None":
        patient_full_title += f" {academic_title}"

    # Hardcoded logo folder path - Change this line if the path changes
    logo_folder = r"C:\Program Files\Motionlab Report\Bericht_logos"
    logo_path = os.path.join(logo_folder, "logo_orthopassion.png")
    second_logo_path = os.path.join(logo_folder, "logo_2_orthopassion.png")

    # Determine file paths
    if export_format == "ODT":
        odt_path = save_path
        pdf_path = None
    elif export_format == "PDF":
        # Create ODT in temp, then convert to PDF
        odt_path = save_path.replace(".pdf", ".odt")
        if odt_path == save_path:
            odt_path = save_path + ".odt"
        pdf_path = save_path
    else:  # BOTH
        if save_path.endswith(".pdf"):
            pdf_path = save_path
            odt_path = save_path.replace(".pdf", ".odt")
        else:
            odt_path = save_path
            pdf_path = save_path.replace(".odt", ".pdf")

    try:
        create_report(patient_full_title, patient_name, patient_dob, report_creator, odt_path, gender, ini_path,
                      sim_performed, isg_right, isg_left, markers, logo_path, second_logo_path,
                      screenshot_path, statik_screenshots, pelvic_drop_sentence, gehen_screenshots, kraft_screenshots,
                      beinachsen_texts, ganganalyse_texts, therapie_texts, measurement_type, ios_pedografie_screenshots,
                      measurement_date, leg_length_texts, vgl_screenshot, strength_test_type)

        # Convert to PDF if needed
        if export_format in ["PDF", "BOTH"]:
            try:
                # Find LibreOffice executable
                libreoffice_path = find_libreoffice()
                if libreoffice_path is None:
                    raise FileNotFoundError(
                        "LibreOffice not found. Please install LibreOffice from https://www.libreoffice.org/download/download/"
                    )

                # Use LibreOffice to convert ODT to PDF
                subprocess.run([
                    libreoffice_path, "--headless", "--convert-to", "pdf",
                    "--outdir", os.path.dirname(pdf_path) or ".",
                    odt_path
                ], check=True)

                # Rename if needed (LibreOffice uses original filename)
                converted_pdf = os.path.join(
                    os.path.dirname(pdf_path) or ".",
                    os.path.basename(odt_path).replace(".odt", ".pdf")
                )
                if converted_pdf != pdf_path and os.path.exists(converted_pdf):
                    os.rename(converted_pdf, pdf_path)

                print(f"PDF created: {pdf_path}")

                # Remove ODT if only PDF was requested
                if export_format == "PDF" and os.path.exists(odt_path):
                    os.remove(odt_path)
                    print(f"Removed temporary ODT: {odt_path}")
            except Exception as e:
                print(f"Error converting to PDF: {e}")
                messagebox.showwarning("PDF Conversion", f"PDF could not be created: {e}\nODT has been saved.")

        # Show success message
        if export_format == "PDF":
            messagebox.showinfo("Success", f"Report created:\n{pdf_path}")
        elif export_format == "ODT":
            messagebox.showinfo("Success", f"Report created:\n{odt_path}")
        else:
            messagebox.showinfo("Success", f"Reports created:\n{pdf_path}\n{odt_path}")

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
# Color scheme from logo
COLOR_TURQUOISE = "#80afaa"
COLOR_BROWN = "#afa190"
COLOR_BG = "#F5F5F5"
COLOR_WHITE = "#FFFFFF"
COLOR_TEXT = "#333333"
COLOR_GREEN = "#4CAF50"
COLOR_RED = "#E57373"


def apply_dialog_style(dialog, title, min_width=400, min_height=300):
    """Apply consistent styling to a Toplevel dialog"""
    dialog.title(title)
    dialog.configure(bg=COLOR_BG)
    dialog.minsize(min_width, min_height)

    # Center on screen
    dialog.update_idletasks()
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    window_width = max(dialog.winfo_width(), min_width)
    window_height = max(dialog.winfo_height(), min_height)
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")


def create_styled_label(parent, text, font_size=11, bold=False, fg=None):
    """Create a styled label"""
    font_weight = "bold" if bold else "normal"
    return tk.Label(
        parent,
        text=text,
        font=("Helvetica", font_size, font_weight),
        fg=fg or COLOR_TEXT,
        bg=COLOR_BG
    )


def create_styled_button(parent, text, command, primary=True, width=12):
    """Create a styled button"""
    bg = COLOR_TURQUOISE if primary else COLOR_BROWN
    active_bg = "#6d9994" if primary else "#9e9180"
    return tk.Button(
        parent,
        text=text,
        command=command,
        font=("Helvetica", 10),
        width=width,
        bg=bg,
        fg=COLOR_WHITE,
        activebackground=active_bg,
        activeforeground=COLOR_WHITE,
        relief=tk.FLAT,
        cursor="hand2"
    )


def create_styled_radiobutton(parent, text, variable, value):
    """Create a styled radiobutton"""
    return tk.Radiobutton(
        parent,
        text=text,
        variable=variable,
        value=value,
        font=("Helvetica", 11),
        bg=COLOR_BG,
        fg=COLOR_TEXT,
        activebackground=COLOR_BG,
        selectcolor=COLOR_WHITE
    )


def create_styled_checkbutton(parent, text, variable):
    """Create a styled checkbutton"""
    return tk.Checkbutton(
        parent,
        text=text,
        variable=variable,
        font=("Helvetica", 11),
        bg=COLOR_BG,
        fg=COLOR_TEXT,
        activebackground=COLOR_BG,
        selectcolor=COLOR_WHITE
    )


def create_styled_frame(parent):
    """Create a styled frame"""
    return tk.Frame(parent, bg=COLOR_BG)


def create_styled_entry(parent, width=40):
    """Create a styled entry"""
    return tk.Entry(
        parent,
        width=width,
        font=("Helvetica", 11),
        relief=tk.SOLID,
        bd=1
    )


def show_needed_files_dialog():
    """Show a non-modal dialog with required files for each report type"""
    dialog = tk.Toplevel(root)
    dialog.title("Needed Files for each Report Type")
    dialog.configure(bg=COLOR_BG)
    dialog.geometry("550x580")

    # Center on screen
    dialog.update_idletasks()
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (275)
    y = (screen_height // 2) - (290)
    dialog.geometry(f"550x580+{x}+{y}")

    # Main container
    main_frame = tk.Frame(dialog, bg=COLOR_BG)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

    # Title
    title_label = tk.Label(
        main_frame,
        text="Required Files by Report Type",
        font=("Helvetica", 14, "bold"),
        fg=COLOR_TEXT,
        bg=COLOR_BG
    )
    title_label.pack(pady=(0, 5))

    # Teal-brown separator line
    sep_frame = tk.Frame(main_frame, bg=COLOR_BG, height=3)
    sep_frame.pack(fill=tk.X, pady=(5, 15))
    tk.Frame(sep_frame, bg=COLOR_TURQUOISE, height=3, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Frame(sep_frame, bg=COLOR_BROWN, height=3, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)

    # Scrollable content
    canvas = tk.Canvas(main_frame, bg=COLOR_BG, highlightthickness=0)
    scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=COLOR_BG)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # File requirements data
    file_requirements = {
        "IOS": [
            "4d_average.pdf  (fallback: statik.pdf)",
            "statik.pdf",
            "ios.pdf",
            ".ini file containing '4D average'",
            "vgl.pdf  (optional - for leg length examination)"
        ],
        "Statik": [
            "4d_average.pdf  (fallback: statik.pdf)",
            "statik.pdf",
            "ios.pdf",
            "kraft.pdf",
            ".ini file containing '4D average'",
            "vgl.pdf  (optional - for leg length examination)"
        ],
        "Walking (Gehen)": [
            "4d_average.pdf  (fallback: statik.pdf)",
            "statik.pdf",
            "gehen.pdf",
            "kraft.pdf",
            ".ini file containing '4D average'",
            ".ini file containing '4D motion'",
            "vgl.pdf  (optional - for leg length examination)"
        ],
        "Running (Laufen)": [
            "4d_average.pdf  (fallback: statik.pdf)",
            "statik.pdf",
            "hp.pdf",
            "ios.pdf",
            "kraft.pdf",
            ".ini file containing '4D average'",
            ".ini file containing '4D motion'",
            "vgl.pdf  (optional - for leg length examination)"
        ]
    }

    for report_type, files in file_requirements.items():
        # Report type header
        type_label = tk.Label(
            scrollable_frame,
            text=report_type,
            font=("Helvetica", 12, "bold"),
            fg=COLOR_BROWN,
            bg=COLOR_BG,
            anchor="w"
        )
        type_label.pack(fill="x", pady=(10, 5))

        # File list
        for file_name in files:
            file_label = tk.Label(
                scrollable_frame,
                text=f"  • {file_name}",
                font=("Helvetica", 10),
                fg=COLOR_TEXT,
                bg=COLOR_BG,
                anchor="w"
            )
            file_label.pack(fill="x", padx=(10, 0))

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Enable mouse wheel scrolling (cross-platform)
    def _on_mousewheel(event):
        if event.num == 4:
            canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            canvas.yview_scroll(3, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel)
    canvas.bind_all("<Button-5>", _on_mousewheel)
    def _on_destroy(event):
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")
    dialog.bind("<Destroy>", _on_destroy)

    # Close button
    close_btn = tk.Button(
        dialog,
        text="Close",
        command=dialog.destroy,
        font=("Helvetica", 10),
        width=12,
        bg=COLOR_RED,
        fg=COLOR_TEXT,
        activebackground="#EF5350",
        activeforeground=COLOR_TEXT,
        relief=tk.FLAT,
        cursor="hand2"
    )
    close_btn.pack(pady=15)


def show_release_notes():
    """Show release notes dialog"""
    dialog = tk.Toplevel(root)
    dialog.title("Release Notes")
    dialog.configure(bg=COLOR_BG)
    dialog.transient(root)
    dialog.grab_set()

    # Center the dialog
    dialog.update_idletasks()
    dialog_width = 420
    dialog_height = 400
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (dialog_width // 2)
    y = (screen_height // 2) - (dialog_height // 2)
    dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
    dialog.minsize(dialog_width, dialog_height)

    # Title
    title_label = tk.Label(
        dialog,
        text="Release Notes - Version 1.5",
        font=("Helvetica", 14, "bold"),
        fg=COLOR_TURQUOISE,
        bg=COLOR_BG
    )
    title_label.pack(pady=(20, 10))

    # Separator
    sep_frame = tk.Frame(dialog, bg=COLOR_BG, height=2)
    sep_frame.pack(fill=tk.X, padx=30, pady=(0, 15))
    tk.Frame(sep_frame, bg=COLOR_TURQUOISE, height=2).pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Frame(sep_frame, bg=COLOR_BROWN, height=2).pack(side=tk.RIGHT, fill=tk.X, expand=True)

    # Release notes content (scrollable)
    scroll_container = tk.Frame(dialog, bg=COLOR_BG)
    scroll_container.pack(fill=tk.BOTH, expand=True, padx=30)

    canvas = tk.Canvas(scroll_container, bg=COLOR_BG, highlightthickness=0)
    scrollbar = tk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)
    notes_frame = tk.Frame(canvas, bg=COLOR_BG)

    notes_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=notes_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Enable mouse wheel scrolling (cross-platform)
    def _on_mousewheel(event):
        if event.num == 4:
            canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            canvas.yview_scroll(3, "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    canvas.bind_all("<Button-4>", _on_mousewheel)
    canvas.bind_all("<Button-5>", _on_mousewheel)
    def _on_destroy(event):
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")
    dialog.bind("<Destroy>", _on_destroy)

    notes = [
        "Version 1.5:",
        "- Added 'Legs' and 'Legs + shoulders' strength test options",
        "- Improved bullet point spacing for better readability",
        "- Updated screenshot coordinates for static and gait/running analysis",
        "- Fixed layout so the leg axis & posture section fits on one page",
        "- Renamed report file to 'YYYY-MM-DD Motionlab Report <name>' format",
        "",
        "Version 1.4:",
        "- Added selection for different strength test types",
        "- Renamed report file from 'ml' to 'motionlab report'",
        "- Included online training info text at the end",
        "",
        "Version 1.3:",
        "- Leg length measurement section added",
        "- Hotfix of PDF export",
        "",
        "Version 1.2:",
        "- Fixed kyphosis/lordosis calculation",
        "- Fixed PDF export on Windows"
    ]

    for note in notes:
        note_label = tk.Label(
            notes_frame,
            text=note,
            font=("Helvetica", 11),
            fg=COLOR_TEXT,
            bg=COLOR_BG,
            anchor="w"
        )
        note_label.pack(fill=tk.X, pady=5)

    # Close button
    close_btn = tk.Button(
        dialog,
        text="Close",
        command=dialog.destroy,
        font=("Helvetica", 10),
        width=12,
        bg=COLOR_TURQUOISE,
        fg=COLOR_WHITE,
        activebackground="#6d9994",
        activeforeground=COLOR_WHITE,
        relief=tk.FLAT,
        cursor="hand2"
    )
    close_btn.pack(pady=20)


root = tk.Tk()
root.title("Motionlab Report Creator")

# Set window size and center on screen
window_width = 550
window_height = 580
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int((screen_width - window_width) / 2)
center_y = int((screen_height - window_height) / 2)
root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
root.configure(bg=COLOR_BG)
root.resizable(True, True)
root.minsize(400, 450)

# Main container frame
main_frame = tk.Frame(root, bg=COLOR_BG)
main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

# Header section with title
header_frame = tk.Frame(main_frame, bg=COLOR_BG)
header_frame.pack(fill=tk.X, pady=(0, 20))

title_label = tk.Label(
    header_frame,
    text="Motionlab Report Creator",
    font=("Helvetica", 22, "bold"),
    fg=COLOR_TEXT,
    bg=COLOR_BG
)
title_label.pack()

# Subtitle/Version
version_label = tk.Label(
    header_frame,
    text="Version 1.5",
    font=("Helvetica", 10),
    fg=COLOR_BROWN,
    bg=COLOR_BG
)
version_label.pack(pady=(5, 0))

# Separator line (left teal, right brown)
sep_frame = tk.Frame(main_frame, bg=COLOR_BG, height=2)
sep_frame.pack(fill=tk.X, pady=(0, 30))
tk.Frame(sep_frame, bg=COLOR_TURQUOISE, height=2, width=1).pack(side=tk.LEFT, fill=tk.X, expand=True)
tk.Frame(sep_frame, bg=COLOR_BROWN, height=2, width=1).pack(side=tk.RIGHT, fill=tk.X, expand=True)

# Description
desc_label = tk.Label(
    main_frame,
    text='Press "Create Report" to start analysis',
    font=("Helvetica", 11),
    fg=COLOR_TEXT,
    bg=COLOR_BG
)
desc_label.pack(pady=(0, 30))

# Buttons frame
button_frame = tk.Frame(main_frame, bg=COLOR_BG)
button_frame.pack(fill=tk.X, pady=10)

# Style for buttons
button_style = {
    "font": ("Helvetica", 11),
    "width": 20,
    "height": 2,
    "cursor": "hand2",
    "relief": tk.FLAT,
    "bd": 0
}

# Create Report button (primary - turquoise)
create_btn = tk.Button(
    button_frame,
    text="Create Report",
    command=generate_report,
    bg=COLOR_TURQUOISE,
    fg=COLOR_WHITE,
    activebackground="#6d9994",
    activeforeground=COLOR_WHITE,
    **button_style
)
create_btn.pack(pady=8)

# Coordinate Finder button (secondary - brown)
coord_btn = tk.Button(
    button_frame,
    text="Coordinate Finder",
    command=find_coordinates,
    bg=COLOR_BROWN,
    fg=COLOR_WHITE,
    activebackground="#9e9180",
    activeforeground=COLOR_WHITE,
    **button_style
)
coord_btn.pack(pady=8)

# Needed Files button (info - lighter style)
files_btn = tk.Button(
    button_frame,
    text="Needed Files for each Report Type",
    command=show_needed_files_dialog,
    bg=COLOR_WHITE,
    fg=COLOR_TEXT,
    activebackground=COLOR_BG,
    activeforeground=COLOR_TEXT,
    font=("Helvetica", 10),
    width=28,
    height=2,
    cursor="hand2",
    relief=tk.SOLID,
    bd=1
)
files_btn.pack(pady=8)

# Release Notes button
release_btn = tk.Button(
    button_frame,
    text="Release Notes",
    command=show_release_notes,
    bg=COLOR_WHITE,
    fg=COLOR_TEXT,
    activebackground=COLOR_BG,
    activeforeground=COLOR_TEXT,
    font=("Helvetica", 10),
    width=28,
    height=2,
    cursor="hand2",
    relief=tk.SOLID,
    bd=1
)
release_btn.pack(pady=8)

# Footer with developer credit
footer_frame = tk.Frame(main_frame, bg=COLOR_BG)
footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))

footer_separator = tk.Frame(footer_frame, height=1, bg=COLOR_BROWN)
footer_separator.pack(fill=tk.X, pady=(0, 10))

developer_label = tk.Label(
    footer_frame,
    text="Developed by Carlo Lade",
    font=("Helvetica", 9),
    fg=COLOR_BROWN,
    bg=COLOR_BG
)
developer_label.pack()

# Clinic name with colored O's
clinic_frame = tk.Frame(footer_frame, bg=COLOR_BG)
clinic_frame.pack(pady=(5, 0))

# "orthopassion - Privatpraxis für regenerative Orthopädie und Osteopathie"
# First 'o' in orthopassion = teal, second 'o' in orthopassion = brown
# "Orthopädie und Osteopathie" in brown/tan color to match website
clinic_parts = [
    ("o", COLOR_TURQUOISE), ("rth", COLOR_TEXT), ("o", COLOR_BROWN), ("passion - Privatpraxis für regenerative ", COLOR_TEXT),
    ("O", COLOR_TURQUOISE), ("rthopädie ", COLOR_TURQUOISE), ("und ", COLOR_TEXT),
    ("O", COLOR_BROWN), ("steopathie", COLOR_BROWN)
]

for text, color in clinic_parts:
    lbl = tk.Label(clinic_frame, text=text, font=("Helvetica", 10), fg=color, bg=COLOR_BG,
                   borderwidth=0, highlightthickness=0, padx=0, pady=0)
    lbl.pack(side=tk.LEFT, padx=0, ipadx=0)

root.mainloop()
