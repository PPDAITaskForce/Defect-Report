import os
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, Side

def apply_box_border(ws, start_row, start_col, end_row, end_col, border_style):
    """Applies outer and inner borders across a range of cells, ensuring merged cells render correctly."""
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).border = border_style

def fit_image_aspect_ratio(img, max_width=430, max_height=220):
    """Resizes openpyxl Image maintaining original aspect ratio."""
    orig_w, orig_h = img.width, img.height
    if orig_w <= 0 or orig_h <= 0:
        img.width, img.height = max_width, max_height
        return

    ratio = min(max_width / orig_w, max_height / orig_h)
    img.width = int(orig_w * ratio)
    img.height = int(orig_h * ratio)

def generate_multi_defect_report(template_path, output_path, header_data, defect_list):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 1. Fill Header Information
    ws['F3'] = header_data.get('project', '')
    ws['F5'] = header_data.get('form_ref', '')
    ws['F7'] = header_data.get('location', '')
    ws['D9'] = header_data.get('date', '')
    ws['D11'] = header_data.get('inspector', '')

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    BLOCK_HEIGHT = 14  # Each block spans 14 rows (e.g., 13 to 26)

    # 2. Process Defect List
    for idx, defect in enumerate(defect_list):
        block_start = 13 + (idx * BLOCK_HEIGHT)

        # Dynamic template generation for blocks >= 3 (index >= 3)
        if idx >= 3:
            # Header Row
            ws.cell(row=block_start, column=2, value="Defects (Before)").font = Font(bold=True)
            ws.cell(row=block_start, column=8, value="Rectification (After)").font = Font(bold=True)
            ws.merge_cells(start_row=block_start, start_column=2, end_row=block_start, end_column=7)
            ws.merge_cells(start_row=block_start, start_column=8, end_row=block_start, end_column=13)

            # Photo Box Borders (Rows +1 to +10)
            apply_box_border(ws, block_start + 1, 2, block_start + 10, 7, thin_border)
            apply_box_border(ws, block_start + 1, 8, block_start + 10, 13, thin_border)

            # Description Labels (Row +11)
            desc_label_row = block_start + 11
            ws.cell(row=desc_label_row, column=2, value="Description").font = Font(bold=True)
            ws.cell(row=desc_label_row, column=8, value="Description").font = Font(bold=True)
            ws.merge_cells(start_row=desc_label_row, start_column=2, end_row=desc_label_row, end_column=7)
            ws.merge_cells(start_row=desc_label_row, start_column=8, end_row=desc_label_row, end_column=13)

            # Description Input Boxes (Rows +12 to +13)
            desc_start = block_start + 12
            desc_end = block_start + 13
            ws.merge_cells(start_row=desc_start, start_column=2, end_row=desc_end, end_column=7)
            ws.merge_cells(start_row=desc_start, start_column=8, end_row=desc_end, end_column=13)
            apply_box_border(ws, desc_start, 2, desc_end, 7, thin_border)
            apply_box_border(ws, desc_start, 8, desc_end, 13, thin_border)

        # Populate Description Text
        desc_cell = ws.cell(row=block_start + 12, column=2)
        desc_cell.value = defect.get('desc', '')
        desc_cell.alignment = Alignment(vertical='top', wrap_text=True)

        # Insert Photo
        img_path = defect.get('image_path')
        if img_path and os.path.exists(img_path):
            img = Image(img_path)
            fit_image_aspect_ratio(img, max_width=430, max_height=220)
            img.anchor = f"B{block_start + 1}"
            ws.add_image(img)

    # Print setup
    ws.print_title_rows = '$1:$12'

    wb.save(output_path)
    print(f"Successfully generated report with {len(defect_list)} entries at: {output_path}")