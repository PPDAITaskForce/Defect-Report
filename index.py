import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import os

def generate_multi_defect_report(template_path, output_path, header_data, defect_list):
    """
    header_data: dict with project, form_ref, location, date, inspector
    defect_list: list of dicts [{'image_path': 'path.png', 'desc': 'text'}, ...]
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 1. Fill Header Info (Rows 1 to 12)
    ws['F3'] = header_data.get('project', 'LOHAS P12')
    ws['F5'] = header_data.get('form_ref', 'P12 - 001')
    ws['F7'] = header_data.get('location', 'Tower 1A 56D')
    ws['D9'] = header_data.get('date', '2026-10-04')
    ws['D11'] = header_data.get('inspector', 'Brian Lau')

    # Border definitions for newly copied defect blocks
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # 2. Process each defect record
    for idx, defect in enumerate(defect_list):
        # Calculate row offset for current defect index
        # Defect 0 -> Block 1 starting row 13
        # Defect 1 -> Block 2 starting row 26
        # Defect 2 -> Block 3 starting row 39
        # Defect N -> Block N+1 starting row 13 + (N * 13)
        block_start = 13 + (idx * 13)

        # If defect index >= 3, construct block structure dynamically
        if idx >= 3:
            # Header label for Defect / Rectification
            ws.cell(row=block_start, column=2, value="Defects (Before)").font = Font(bold=True)
            ws.cell(row=block_start, column=8, value="Rectification (After)").font = Font(bold=True)
            ws.merge_cells(start_row=block_start, start_column=2, end_row=block_start, end_column=7)
            ws.merge_cells(start_row=block_start, start_column=8, end_row=block_start, end_column=13)

            # Description Label Row
            desc_label_row = block_start + 11
            ws.cell(row=desc_label_row, column=2, value="Description").font = Font(bold=True)
            ws.cell(row=desc_label_row, column=8, value="Description ").font = Font(bold=True)

            # Description Text Box (Merged 2 rows x 6 columns)
            desc_box_start = block_start + 12
            desc_box_end = block_start + 13
            ws.merge_cells(start_row=desc_box_start, start_column=2, end_row=desc_box_end, end_column=7)
            ws.merge_cells(start_row=desc_box_start, start_column=8, end_row=desc_box_end, end_column=13)

            # Draw photo container borders B(col 2) to G(col 7)
            photo_start_row = block_start + 1
            photo_end_row = block_start + 10
            for r in range(photo_start_row, photo_end_row + 1):
                for c in range(2, 8):
                    ws.cell(row=r, column=c).border = thin_border

            # Set row height for photo box base
            ws.row_dimensions[photo_end_row].height = 15.0

        # Fill Description text
        desc_cell = ws.cell(row=block_start + 12, column=2)
        desc_cell.value = defect.get('desc', '')
        desc_cell.alignment = Alignment(vertical='top', wrap_text=True)

        # Insert Photo into Columns B to G
        img_path = defect.get('image_path')
        if img_path and os.path.exists(img_path):
            img = Image(img_path)
            # Scale image to fit within B to G (width ~430px) and 10 rows height (~220px)
            img.width = 430
            img.height = 220
            
            # Anchor image to top-left cell of the photo box (Cell B14, B27, B40, etc.)
            photo_cell_anchor = f"B{block_start + 1}"
            img.anchor = photo_cell_anchor
            ws.add_image(img)

    # 3. Repeat Header rows 1 to 12 when printing multi-page Excel files
    ws.print_title_rows = '$1:$12'

    # Save finalized report
    wb.save(output_path)
    print(f"Report successfully generated with {len(defect_list)} defects at: {output_path}")


# --- TEST EXECUTION ---
if __name__ == "__main__":
    header_info = {
        'project': 'LOHAS P12',
        'form_ref': 'P12 - 001',
        'location': 'Tower 1A 56D',
        'date': '2026-10-04',
        'inspector': 'Brian Lau'
    }

    # Example with 5 defect records
    sample_defects = [
        {'image_path': 'photo1.jpg', 'desc': 'Defect 1: Incompleted A/C Platform'},
        {'image_path': 'photo2.jpg', 'desc': 'Defect 2: Wall paint peel near balcony'},
        {'image_path': 'photo3.jpg', 'desc': 'Defect 3: Window frame scratch'},
        {'image_path': 'photo4.jpg', 'desc': 'Defect 4: Tiles misalignment'},
        {'image_path': 'photo5.jpg', 'desc': 'Defect 5: Water seepage under kitchen sink'}
    ]

    generate_multi_defect_report("PPD-016.xlsx", "PPD-016_Extended_Report.xlsx", header_info, sample_defects)