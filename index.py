import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import os


def generate_defect_report(template_path, output_path, image_path, report_data):
    # Load the existing workbook template
    if os.path.exists(template_path):
        wb = openpyxl.load_workbook(template_path)
        ws = wb.active
    else:
        # Fallback if template file path doesn't exist directly
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Defect Report"

        # Basic Header setup matching PPD-016 template standard
        ws.merge_cells("B2:G2")
        ws["B2"] = "DEFECT REPORT (PPD-016)"
        ws["B2"].font = Font(size=16, bold=True)
        ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    # 1. Fill basic metadata into the sheet
    # Adjust cell references below to match your template fields
    ws["C4"] = report_data.get("report_no", "DEF-001")
    ws["C5"] = report_data.get("date", "2026-10-04")
    ws["C6"] = report_data.get("location", "Site A")
    ws["F4"] = report_data.get("inspector", "John Doe")
    ws["F5"] = report_data.get("severity", "High")

    # Defect Description (Merged B8:G8)
    ws["B8"] = f"Description: {report_data.get('description', 'N/A')}"

    # 2. Embed Defect Photo into Columns B to G
    if os.path.exists(image_path):
        img = Image(image_path)

        # Calculate pixel width/height to fit Columns B through G nicely
        # Column B to G total width is ~ 450-500 pixels; Height (Rows 10-20) ~ 250-300 pixels
        img.width = 460
        img.height = 280

        # Anchor the top-left corner of the image to Cell B10
        img.anchor = "B10"

        # Add image to worksheet
        ws.add_image(img)
    else:
        print(f"Warning: Image file not found at {image_path}")

    # 3. Format and Merge Photo Placeholder Area (B10:G20) for visual outline
    ws.merge_cells("B10:G20")

    # Apply thin borders around the image frame
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for row in ws["B10:G20"]:
        for cell in row:
            cell.border = thin_border

    # Save output file
    wb.save(output_path)
    print(f"Defect report successfully generated and saved to: {output_path}")


# --- Usage Example ---
if __name__ == "__main__":
    template_file = "PPD-016.xlsx"
    output_file = "Generated_Defect_Report.xlsx"
    photo_file = "defect_photo.jpg"  # Replace with your actual photo file path

    data = {
        "report_no": "PPD-016-2026-001",
        "date": "2026-10-04",
        "location": "Block 3 - Level 4",
        "inspector": "Alex Smith",
        "severity": "Major",
        "description": "Concrete spalling detected on structural column B2.",
    }

    generate_defect_report(
        template_file, output_file, photo_file, report_data=data
    )