import json
import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DEFAULT_FIELDS = [
    "TENBENHNHAN",
    "NGAYSINH",
    "GIOITINH",
    "MAHOSOBENHAN",
    "DIACHI",
    "SDTBENHNHAN",
]


def export_patients_to_excel(
    data_source,
    hospital_name: str = "bvlaichau",
    fields: list = None,
    output_dir: str = ".",
) -> dict:
    """
    Xuất danh sách bệnh nhân từ file json hoặc dict/list sang file Excel (.xlsx).
    Tên file tự động định dạng: [tên bệnh viện][số lượng].xlsx

    Args:
        data_source: Đường dẫn tới file JSON, hoặc dict chứa 'rows', hoặc list các dict bệnh nhân.
        hospital_name: Tên bệnh viện (mặc định 'bvlaichau').
        fields: Danh sách các cột cần lấy.
        output_dir: Thư mục lưu file Excel.

    Returns:
        dict: {success, file_path, count, records}
    """
    if fields is None:
        fields = DEFAULT_FIELDS

    # Đọc dữ liệu
    if isinstance(data_source, str):
        if not os.path.exists(data_source):
            return {"success": False, "message": f"Không tìm thấy file: {data_source}"}
        with open(data_source, "r", encoding="utf-8") as f:
            data = json.load(f)
    elif isinstance(data_source, dict):
        data = data_source
    elif isinstance(data_source, list):
        data = {"rows": data_source}
    else:
        return {"success": False, "message": "Định dạng dữ liệu không hợp lệ"}

    rows = data.get("rows", [])
    count = len(rows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Danh sách bệnh nhân"

    # Header
    ws.append(fields)

    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    data_font = Font(name="Arial", size=10)
    for r in rows:
        # Hỗ trợ lấy từ raw dictionary hoặc cleaned dictionary
        row_vals = []
        for field in fields:
            val = r.get(field)
            if val is None and "raw" in r and isinstance(r["raw"], dict):
                val = r["raw"].get(field)
            row_vals.append("" if val is None else str(val))
        ws.append(row_vals)

    # Styling data cells
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(fields)):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            # Căn giữa các cột mã, ngày sinh, giới tính, sđt
            if cell.column in (2, 3, 4, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Tự động căn chỉnh độ rộng cột
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # Tên file duy nhất: [tên bệnh viện][số lượng].xlsx
    filename = f"[{hospital_name}][{count}].xlsx"
    file_path = os.path.join(output_dir, filename)

    wb.save(file_path)

    return {
        "success": True,
        "file_path": file_path,
        "filename": filename,
        "count": count,
        "headers": fields,
    }


if __name__ == "__main__":
    res = export_patients_to_excel("raw.json", hospital_name="bvlaichau")
    print("Export result:", res)
