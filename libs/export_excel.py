import json
import os
import re
import zipfile
import openpyxl

DEFAULT_FIELDS = [
    "KHAMBENHID",
    "MABENHNHAN",
    "TENBENHNHAN",
    "MA_BHYT",
    "NGAYSINH",
    "GIOITINH",
    "ORG_NAME",
    "MAHOSOBENHAN",
    "NOILAMVIEC",
    "BACSY",
    "XUTRIKHAMBENH",
    "MAKCBBD",
    "SDTBENHNHAN",
    "NGAYTIEPNHAN",
    "DIACHI",
    "NGAYRAVIEN",
    "NGUOITN",
    "DOITUONGBENHNHAN",
    "KENHHENKHAM",
    "KHOA_NV",
    "SOVAOVIEN",
    "CHANDOANRAVIEN",
    "TUYENDK",
    "DANTOC",
    "NAMSINH",
    "SDT_LIENHE",
]


def _extract_field(r: dict, field: str):
    """
    Trích xuất giá trị trường dữ liệu bệnh nhân từ record (hỗ trợ cả raw dict và clean dict),
    tự động suy luận các trường theo chuẩn biểu mẫu Mauu.xlsx.
    """
    raw = r.get("raw") if isinstance(r.get("raw"), dict) else r
    val = r.get(field)
    if val is None or str(val).strip() == "":
        val = raw.get(field)

    # Suy luận / ánh xạ thông minh cho các trường đặc thù trong biểu mẫu Mauu.xlsx
    if val is None or str(val).strip() == "":
        if field == "KENHHENKHAM":
            val = r.get("KENHHENKHAM") or raw.get("KENHHENKHAM") or "Tại quầy"
        elif field == "DANTOC":
            val = r.get("DANTOC") or raw.get("DANTOC") or r.get("dan_toc") or "Kinh"
        elif field == "NAMSINH":
            val = r.get("NAMSINH") or raw.get("NAMSINH")
            if not val:
                ngaysinh = str(r.get("NGAYSINH") or raw.get("NGAYSINH") or r.get("ngay_sinh") or "")
                m = re.search(r"\b(19\d{2}|20\d{2})\b", ngaysinh)
                val = m.group(1) if m else None
        elif field == "SDT_LIENHE":
            val = (
                r.get("SDT_LIENHE")
                or raw.get("SDT_LIENHE")
                or r.get("SDTBENHNHAN")
                or raw.get("SDTBENHNHAN")
                or r.get("sdt")
                or None
            )
        elif field == "TUYENDK":
            val = r.get("TUYENDK") or raw.get("TUYENDK")
            if not val:
                tuyen = str(r.get("TUYEN") or raw.get("TUYEN") or "").strip()
                if tuyen:
                    val = tuyen
                elif r.get("DUNGTUYEN") in ("x", "1") or raw.get("DUNGTUYEN") in ("x", "1"):
                    val = "Đúng tuyến"
                elif r.get("TRAITUYEN") in ("x", "1") or raw.get("TRAITUYEN") in ("x", "1"):
                    val = "Trái tuyến"
                else:
                    val = None
        elif field == "SOVAOVIEN":
            val = r.get("SOVAOVIEN") or raw.get("SOVAOVIEN") or r.get("SOTHUTU") or raw.get("SOTHUTU") or None
        elif field == "CHANDOANRAVIEN":
            val = (
                r.get("CHANDOANRAVIEN")
                or raw.get("CHANDOANRAVIEN")
                or r.get("CHANDOAN")
                or raw.get("CHANDOAN")
                or r.get("CONGKHAM")
                or raw.get("CONGKHAM")
                or r.get("cong_kham")
                or None
            )
        elif field == "KHAMBENHID":
            val = r.get("KHAMBENHID") or raw.get("KHAMBENHID") or r.get("kham_benh_id")
        elif field == "MABENHNHAN":
            val = r.get("MABENHNHAN") or raw.get("MABENHNHAN") or r.get("ma_bn")
        elif field == "TENBENHNHAN":
            val = r.get("TENBENHNHAN") or raw.get("TENBENHNHAN") or r.get("ten_bn")
        elif field == "MA_BHYT":
            val = r.get("MA_BHYT") or raw.get("MA_BHYT") or r.get("ma_bhyt")
        elif field == "NGAYSINH":
            val = r.get("NGAYSINH") or raw.get("NGAYSINH") or r.get("ngay_sinh")
        elif field == "GIOITINH":
            val = r.get("GIOITINH") or raw.get("GIOITINH") or r.get("gioi_tinh")
        elif field == "ORG_NAME":
            val = r.get("ORG_NAME") or raw.get("ORG_NAME") or r.get("phong_kham")
        elif field == "MAHOSOBENHAN":
            val = r.get("MAHOSOBENHAN") or raw.get("MAHOSOBENHAN") or r.get("ma_ho_so")
        elif field == "SDTBENHNHAN":
            val = r.get("SDTBENHNHAN") or raw.get("SDTBENHNHAN") or r.get("sdt")
        elif field == "NGAYTIEPNHAN":
            val = r.get("NGAYTIEPNHAN") or raw.get("NGAYTIEPNHAN") or r.get("ngay_tiep_nhan")
        elif field == "DIACHI":
            val = r.get("DIACHI") or raw.get("DIACHI") or r.get("dia_chi")
        elif field == "DOITUONGBENHNHAN":
            val = r.get("DOITUONGBENHNHAN") or raw.get("DOITUONGBENHNHAN") or r.get("doi_tuong")
        elif field == "BACSY":
            val = r.get("BACSY") or raw.get("BACSY") or r.get("bac_sy")

    if val is None or str(val).strip() == "":
        return None
    return str(val).strip()


def export_patients_to_excel(
    data_source,
    hospital_name: str = "bvlaichau",
    fields: list = None,
    output_dir: str = ".",
    auto_zip_threshold_mb: float = 5.0,
) -> dict:
    """
    Xuất danh sách bệnh nhân từ file json hoặc dict/list sang file Excel (.xlsx)
    chuẩn hóa 100% theo mẫu Mauu.xlsx với tốc độ cao (write_only mode streaming):
    - Sheet title: 'Sheet1'
    - Dòng 1: Để trống
    - Dòng 2: Header 26 cột chuẩn
    - Dòng 3+: Dữ liệu bệnh nhân
    - Tên file tự động định dạng: [tên bệnh viện][số lượng].xlsx
    - Tự động nén sang .zip nếu kích thước file > auto_zip_threshold_mb (mặc định 5MB)
      giúp upload siêu tốc lên Telegram mà không bị nghẽn mạng VPS.

    Args:
        data_source: Đường dẫn tới file JSON, hoặc dict chứa 'rows', hoặc list các dict bệnh nhân.
        hospital_name: Tên bệnh viện.
        fields: Danh sách các cột cần lấy (mặc định lấy 26 cột chuẩn Mauu.xlsx).
        output_dir: Thư mục lưu file Excel.
        auto_zip_threshold_mb: Ngưỡng dung lượng (MB) để tự động nén .zip (0 để tắt).

    Returns:
        dict: {success, file_path, raw_file_path, filename, is_zip, original_size, file_size, count, headers}
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

    # Sử dụng write_only=True để ghi dữ liệu dạng stream cực nhanh, không tốn RAM khi dữ liệu lên tới hàng trăm nghìn dòng
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title="Sheet1")

    # Căn chỉnh độ rộng cột theo mẫu Mauu.xlsx (chuẩn 13.0)
    for col_idx in range(1, len(fields) + 1):
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 13.0

    # Dòng 1: Để trống theo mẫu Mauu.xlsx
    ws.append([])

    # Dòng 2: Header theo mẫu Mauu.xlsx
    ws.append(fields)

    # Dòng 3+: Dữ liệu bệnh nhân (streaming trực tiếp vào zip stream)
    for r in rows:
        row_vals = [_extract_field(r, field) for field in fields]
        ws.append(row_vals)

    # Tên file duy nhất: [tên bệnh viện][số lượng].xlsx
    filename = f"[{hospital_name}][{count}].xlsx"
    file_path = os.path.join(output_dir, filename)

    wb.save(file_path)
    xlsx_size = os.path.getsize(file_path)

    # Tự động nén sang .zip nếu dung lượng vượt ngưỡng (tiết kiệm băng thông và tăng tốc độ tải lên Telegram)
    if auto_zip_threshold_mb and auto_zip_threshold_mb > 0 and xlsx_size > auto_zip_threshold_mb * 1024 * 1024:
        zip_filename = f"[{hospital_name}][{count}].zip"
        zip_path = os.path.join(output_dir, zip_filename)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            z.write(file_path, arcname=filename)
        zip_size = os.path.getsize(zip_path)
        return {
            "success": True,
            "file_path": zip_path,
            "raw_file_path": file_path,
            "filename": zip_filename,
            "is_zip": True,
            "original_size": xlsx_size,
            "file_size": zip_size,
            "count": count,
            "headers": fields,
        }

    return {
        "success": True,
        "file_path": file_path,
        "raw_file_path": file_path,
        "filename": filename,
        "is_zip": False,
        "original_size": xlsx_size,
        "file_size": xlsx_size,
        "count": count,
        "headers": fields,
    }


if __name__ == "__main__":
    res = export_patients_to_excel("raw.json", hospital_name="bvlaichau")
    print("Export result:", res)
