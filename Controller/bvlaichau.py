import asyncio
from datetime import datetime
import json
import logging
import os
import re
import sys
import time
from urllib.parse import urljoin
import httpx

# Thêm project root vào sys.path để hỗ trợ chạy trực tiếp
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from libs.login import login
from libs.sel_dept import sel_dept
from libs.export_excel import export_patients_to_excel, DEFAULT_FIELDS
from libs.get_menu import get_reception_menu

log = logging.getLogger(__name__)

REQUIRED_PATIENT_FIELDS = [
    "TENBENHNHAN",
    "NGAYSINH",
    "GIOITINH",
    "MAHOSOBENHAN",
    "DIACHI",
    "SDTBENHNHAN",
]


def has_required_patient_fields(row: dict) -> bool:
    """Kiểm tra xem bản ghi trả về từ procedure có chứa các trường thông tin bệnh nhân hay không."""
    if not isinstance(row, dict):
        return False
    keys = set(row.keys())
    # Kiểm tra các trường quan trọng nhất của bệnh nhân
    return all(f in keys for f in ["TENBENHNHAN", "NGAYSINH", "GIOITINH", "MAHOSOBENHAN"])

_DEFAULT_URL = "https://bvlaichau.vncare.vn/vnpthis"
_DEFAULT_USER = "TTKSBT.ADMINLCU"
_DEFAULT_PASS = "Lcu@6886"

_HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "vi;q=0.5",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/155.0.0.0 Safari/537.36"
    ),
    "X-Requested-With": "XMLHttpRequest",
}


DATE_OPTIONS = {
    "1": {"label": "Từ 00:00 01/01/2021 đến nay", "start_year": 2021},
    "2": {"label": "Từ 00:00 01/01/2022 đến nay", "start_year": 2022},
    "3": {"label": "Từ 00:00 01/01/2023 đến nay", "start_year": 2023},
    "4": {"label": "Từ 00:00 01/01/2024 đến nay", "start_year": 2024},
    "5": {"label": "Từ 00:00 01/01/2025 đến nay", "start_year": 2025},
    "6": {"label": "Từ 00:00 01/01/2026 đến nay", "start_year": 2026},
    "today": {"label": "Hôm nay", "start_year": None},
}


def resolve_date_range(
    option: str | int = None,
    target_date: str = None,
    from_date: str = None,
    to_date: str = None,
) -> tuple[str, str, str]:
    """
    Xác định khoảng thời gian lọc dữ liệu theo các option hoặc ngày cụ thể.

    Options:
        - "1" hoặc 2021: Từ 00:00 01/01/2021 --> nay
        - "2" hoặc 2022: Từ 00:00 01/01/2022 --> nay
        - "3" hoặc 2023: Từ 00:00 01/01/2023 --> nay
        - "4" hoặc 2024: Từ 00:00 01/01/2024 --> nay
        - "5" hoặc 2025: Từ 00:00 01/01/2025 --> nay
        - "6" hoặc 2026: Từ 00:00 01/01/2026 --> nay
        - "today" hoặc None: Hôm nay

    Returns:
        tuple: (from_date, to_date, label)
    """
    now = datetime.now()
    now_str = now.strftime("%d/%m/%Y")
    default_to_date = now_str

    # Nếu đã chỉ định tường minh cả from_date và to_date
    if from_date and to_date:
        f_d = from_date.split()[0]
        t_d = to_date.split()[0]
        return f_d, t_d, f"{from_date} - {to_date}"

    # Kiểm tra option dạng số năm (2021..2026)
    if str(option) in [str(y) for y in range(2021, 2027)]:
        year = str(option)
        return f"01/01/{year}", default_to_date, f"Từ 01/01/{year} đến nay"

    # Kiểm tra option theo key ("1".."6")
    str_opt = str(option).strip() if option is not None else ""
    if str_opt in DATE_OPTIONS and DATE_OPTIONS[str_opt]["start_year"]:
        year = DATE_OPTIONS[str_opt]["start_year"]
        return f"01/01/{year}", default_to_date, DATE_OPTIONS[str_opt]["label"]

    # Mặc định theo ngày chỉ định hoặc hôm nay
    date_str = target_date if target_date else now_str
    d_clean = date_str.split()[0]
    return d_clean, d_clean, f"Ngày {d_clean}"


def _get_default_date_range(target_date: str = None) -> tuple[str, str]:
    """Helper tương thích ngược."""
    f_date, t_date, _ = resolve_date_range(target_date=target_date)
    return f_date, t_date


def _clean_patient(p: dict) -> dict:
    """Rút gọn và chuẩn hóa các thông tin bệnh nhân quan trọng từ cả v1 (NGT.002) và v2 (NGT.001.BVTM.01)."""
    return {
        "stt": p.get("SOTHUTU") or p.get("RN", ""),
        "ma_bn": p.get("MABENHNHAN", ""),
        "ten_bn": p.get("TENBENHNHAN", ""),
        "gioi_tinh": p.get("GIOITINH", ""),
        "ngay_sinh": p.get("NGAYSINH", ""),
        "ma_bhyt": p.get("MA_BHYT", ""),
        "doi_tuong": p.get("DOITUONGBENHNHAN", ""),
        "dia_chi": p.get("DIACHI", ""),
        "sdt": p.get("SDTBENHNHAN", ""),
        "dan_toc": p.get("DANTOC", ""),
        "vssid": p.get("VSSID", ""),
        "ngay_tiep_nhan": p.get("NGAYTIEPNHAN", ""),
        "phong_kham": p.get("ORG_NAME", ""),
        "cong_kham": p.get("CONGKHAM", ""),
        "ma_ho_so": p.get("MAHOSOBENHAN", ""),
        "bac_sy": p.get("BACSY", ""),
        "kham_benh_id": p.get("KHAMBENHID", ""),
        "trang_thai": p.get("TRANGTHAIKHAMBENH", ""),
        "raw": p,
    }


async def _extract_uuid(client: httpx.AsyncClient, base_url: str, location: str) -> str:
    """Truy cập manager.jsp để trích xuất JWT uuid."""
    manager_url = urljoin(base_url, location) if location else base_url + "/main/manager.jsp"
    r = await client.get(manager_url, headers=_HEADERS, follow_redirects=True)
    match = re.search(r"var\s+uuid\s*=\s*'([^']+)'", r.text)
    if not match:
        match = re.search(r'var\s+uuid\s*=\s*"([^"]+)"', r.text)
    return match.group(1) if match else ""


def _merge_patient_records(list_v1: list, list_v2: list) -> list:
    """
    Hợp nhất và khử trùng lặp danh sách bệnh nhân từ v1 (NGT.002) và v2 (NGT.001.BVTM.01).
    Ưu tiên giữ đầy đủ các trường dữ liệu từ cả 2 phiên bản (không phân tách v1 hay v2).
    """
    merged_map = {}

    for p in list_v1:
        raw = p.get("raw", p) if isinstance(p, dict) else p
        key = str(raw.get("KHAMBENHID") or raw.get("MAHOSOBENHAN") or raw.get("MABENHNHAN") or id(raw))
        merged_map[key] = raw.copy()

    for p in list_v2:
        raw = p.get("raw", p) if isinstance(p, dict) else p
        key = str(raw.get("KHAMBENHID") or raw.get("MAHOSOBENHAN") or raw.get("MABENHNHAN") or id(raw))
        if key in merged_map:
            # Bổ sung các trường có dữ liệu từ v2 sang nếu v1 còn thiếu
            for k, val in raw.items():
                if val is not None and str(val).strip() != "":
                    if not merged_map[key].get(k):
                        merged_map[key][k] = val
        else:
            merged_map[key] = raw.copy()

    return list(merged_map.values())


async def get_patients_by_room(
    client: httpx.AsyncClient,
    base_url: str,
    uuid: str,
    room_id: str = "0",
    date_option: str | int = None,
    from_date: str = None,
    to_date: str = None,
    rows: int = 5000,
    page: int = 1,
    fetch_all: bool = False,
    api_version: str = "v1",
    procedure_name: str = None,
    progress_callback=None,
) -> dict:
    """
    Lấy danh sách bệnh nhân của một phòng cụ thể (hoặc tất cả các phòng nếu room_id='0').

    Args:
        client:         httpx.AsyncClient đang giữ cookies phiên đăng nhập
        base_url:       Base URL bệnh viện (vd: https://bvlaichau.vncare.vn/vnpthis)
        uuid:           JWT token phiên làm việc
        room_id:        ID phòng khám (mặc định '0' = tất cả các phòng)
        procedure_name: Tên Stored Procedure (lấy từ GridUtil.loadGridBySqlPage trong menu, vd: 'NGT.002', 'NGT.001.BVTM.01')
        date_option:    Option chọn khoảng thời gian ('1'..'6', 2021..2026, 'today')
        from_date:      Thời gian bắt đầu (vd: '19/10/2026 00:00:00')
        to_date:        Thời gian kết thúc (vd: '19/10/2026 23:59:59')
        rows:           Số dòng mỗi trang (mặc định 5000)
        page:           Trang bắt đầu
        fetch_all:      Nếu True, sẽ tự động duyệt qua tất cả các trang để lấy đủ toàn bộ dữ liệu.
        api_version:    'v1' (NGT.002) hoặc 'v2' (NGT.001.BVTM.01) nếu không truyền procedure_name

    Returns:
        dict: {room_id, date_label, total_pages, page, records, patients}
    """
    f_date, t_date, date_label = resolve_date_range(
        option=date_option,
        from_date=from_date,
        to_date=to_date,
    )

    if not procedure_name:
        is_v2 = str(api_version).lower() in ("v2", "2", "vs2")
        procedure_name = "NGT.001.BVTM.01" if is_v2 else "NGT.002"
    else:
        is_v2 = (
            procedure_name == "NGT.001.BVTM.01"
            or "BVTM" in procedure_name
            or str(api_version).lower() in ("v2", "2", "vs2")
        )

    referer_func = (
        "../ngoaitru/NGT01T001_tiepnhan_ngt_vs2"
        if is_v2
        else "../ngoaitru/NGT01T001_tiepnhan_ngt"
    )

    f_date = str(f_date).strip().split()[0]
    t_date = str(t_date).strip().split()[0]

    options = [
        {"name": "[0]", "value": f_date},
        {"name": "[1]", "value": t_date},
        {"name": "[3]", "value": "0"},
        {"name": "[4]", "value": str(room_id)},
        {"name": "[5]", "value": "0"},
        {"name": "[6]", "value": "0"},
        {"name": "[7]", "value": "-1"},
    ]
    if is_v2:
        options.append({"name": "[8]", "value": ""})

    post_data_obj = {
        "func": "ajaxExecuteQueryPaging",
        "uuid": uuid,
        "params": [procedure_name],
        "options": options,
    }

    params = {
        "postData": json.dumps(post_data_obj),
        "_search": "false",
        "nd": str(int(time.time() * 1000)),
        "rows": str(rows),
        "page": str(page),
        "sidx": "",
        "sord": "asc",
    }

    rest_url = base_url.rstrip("/") + "/RestService"
    referer_url = f"{base_url.rstrip('/')}/main/manager.jsp?func={referer_func}"

    resp = await client.get(
        rest_url,
        params=params,
        headers={**_HEADERS, "Referer": referer_url},
    )

    try:
        data = resp.json()
    except Exception:
        data = json.loads(resp.text)

    records = int(data.get("records", 0))
    total_pages = int(data.get("total", 1))
    raw_rows = list(data.get("rows", []))

    # Nếu yêu cầu fetch_all và có nhiều trang, tải đồng thời đa luồng (parallel async) tất cả các trang còn lại
    if fetch_all and total_pages > 1:
        sem = asyncio.Semaphore(10)
        completed_pages = [1]

        async def _fetch_single_page(p_idx: int):
            async with sem:
                page_params = dict(params)
                page_params["page"] = str(p_idx)
                page_params["nd"] = str(int(time.time() * 1000) + p_idx)
                try:
                    p_resp = await client.get(
                        rest_url,
                        params=page_params,
                        headers={**_HEADERS, "Referer": referer_url},
                    )
                    try:
                        p_data = p_resp.json()
                    except Exception:
                        p_data = json.loads(p_resp.text)
                    batch = list(p_data.get("rows", []))
                    completed_pages[0] += 1
                    if progress_callback:
                        try:
                            await progress_callback(
                                f"📥 Đang tải dữ liệu: trang {completed_pages[0]}/{total_pages} ({min(completed_pages[0] * rows, records):,}/{records:,} bản ghi)..."
                            )
                        except Exception:
                            pass
                    return p_idx, batch
                except Exception as ex:
                    log.warning(f"Lỗi tải trang {p_idx} procedure {procedure_name}: {ex}")
                    return p_idx, []

        page_tasks = [_fetch_single_page(p) for p in range(2, total_pages + 1)]
        page_results = await asyncio.gather(*page_tasks)
        page_results.sort(key=lambda x: x[0])
        for _, rows_batch in page_results:
            raw_rows.extend(rows_batch)

    patients = [_clean_patient(p) for p in raw_rows]

    return {
        "room_id": str(room_id),
        "date_label": date_label,
        "from_date": f_date,
        "to_date": t_date,
        "total_pages": total_pages,
        "page": int(data.get("page", 1)),
        "records": records,
        "fetched_count": len(patients),
        "patients": patients,
    }


async def get_patient_sample(
    url: str,
    username: str,
    password: str,
    limit: int = 5,
    date_options: tuple = ("today", "6"),
) -> dict:
    """
    Đăng nhập và lấy tối đa `limit` bệnh nhân đầu tiên (v1/v2) để kiểm tra DIACHI.

    Dùng cho luồng lọc tài khoản theo tỉnh — không dựng Excel, không tải toàn bộ.
    Trả về: {success, error_kind, message, url, patients, date_label}
    """
    login_result = await login(url, username, password)
    if not login_result.get("success"):
        return {
            "success": False,
            "error_kind": login_result.get("error_kind", "unexpected"),
            "message": login_result.get("message", "Đăng nhập thất bại"),
            "url": url,
            "patients": [],
        }

    base_url = login_result.get("base_url") or url.rstrip("/")
    cookies = login_result.get("cookies", {})
    location = login_result.get("location", "")

    try:
        async with httpx.AsyncClient(cookies=cookies, verify=False, timeout=30.0) as client:
            uuid = await _extract_uuid(client, base_url, location)
            if not uuid:
                return {"success": False, "error_kind": "unexpected",
                        "message": "Không lấy được token uuid", "url": base_url, "patients": []}

            last_error = "Không lấy được dữ liệu bệnh nhân"
            saw_network = False
            for date_opt in date_options:
                for proc in ("NGT.001.BVTM.01", "NGT.002"):
                    try:
                        res = await get_patients_by_room(
                            client=client,
                            base_url=base_url,
                            uuid=uuid,
                            procedure_name=proc,
                            date_option=date_opt,
                            rows=max(1, int(limit)),
                            fetch_all=False,
                        )
                    except Exception as e:
                        log.debug(f"Sample {proc} ({date_opt}) lỗi trên {base_url}: {e}")
                        last_error = str(e)
                        saw_network = True
                        continue

                    patients = res.get("patients") or []
                    if patients and has_required_patient_fields(patients[0].get("raw", {})):
                        return {
                            "success": True,
                            "error_kind": "ok",
                            "message": "OK",
                            "url": base_url,
                            "date_label": res.get("date_label"),
                            "patients": patients[:limit],
                            "login_result": login_result,
                        }
            # Lỗi mạng khi lấy dữ liệu phải báo là "network" để nơi gọi thử lại
            # thay vì kết luận tài khoản không có bệnh nhân của tỉnh.
            return {"success": False, "error_kind": "network" if saw_network else "unexpected",
                    "message": last_error, "url": base_url, "patients": []}
    except Exception as e:
        return {"success": False, "error_kind": "network",
                "message": str(e), "url": base_url, "patients": []}


async def get_hospital_patients_tree(
    url: str = _DEFAULT_URL,
    username: str = _DEFAULT_USER,
    password: str = _DEFAULT_PASS,
    login_result: dict = None,
    target_date: str = None,
) -> dict:
    """
    Lấy danh sách bệnh nhân của tất cả các khoa và phòng trong tài khoản bệnh viện.

    Args:
        url:          URL bệnh viện
        username:     Tài khoản đăng nhập
        password:     Mật khẩu đăng nhập
        login_result: Kết quả đăng nhập nếu đã có
        target_date:  Ngày cần lấy danh sách (dd/mm/yyyy), mặc định hôm nay

    Returns:
        dict: Cây dữ liệu gồm các khoa, các phòng và danh sách bệnh nhân tương ứng
    """
    # 1. Đăng nhập nếu chưa có login_result
    if not login_result or not login_result.get("success"):
        login_result = await login(url, username, password)
        if not login_result.get("success"):
            return {
                "success": False,
                "message": f"Đăng nhập thất bại: {login_result.get('message')}",
                "departments": [],
                "total_patients": 0,
            }

    cookies = login_result.get("cookies", {})
    location = login_result.get("location", "")
    base_url = login_result.get("base_url") or url.rstrip("/")

    # 2. Lấy danh sách tất cả các khoa và phòng của tài khoản
    dept_res = await sel_dept(url, login_result=login_result)
    departments = dept_res.get("departments", [])

    # Tạo bản đồ ánh xạ từ tên phòng -> thông tin phòng & khoa
    room_map = {}
    for d in departments:
        for r in d.get("rooms", []):
            # Key bằng tên phòng chuẩn hóa (chữ thường, bỏ khoảng trắng thừa)
            norm_name = r["room_name"].strip().lower()
            room_map[norm_name] = {
                "dept_id": d["dept_id"],
                "dept_name": d["dept_name"],
                "room_id": r["room_id"],
                "room_name": r["room_name"],
            }

    from_date, to_date = _get_default_date_range(target_date)

    async with httpx.AsyncClient(
        cookies=cookies,
        verify=False,
        timeout=60.0,
        http1=True,
        http2=False,
    ) as client:
        # 3. Trích xuất uuid
        uuid = await _extract_uuid(client, base_url, location)
        if not uuid:
            return {
                "success": False,
                "message": "Không tìm thấy token uuid phiên làm việc",
                "departments": [],
                "total_patients": 0,
            }

        # 4. Lấy toàn bộ danh sách bệnh nhân tiếp nhận hôm nay (room_id='0')
        all_res = await get_patients_by_room(
            client=client,
            base_url=base_url,
            uuid=uuid,
            room_id="0",
            from_date=from_date,
            to_date=to_date,
            rows=2000,  # Lấy tối đa để bao quát toàn viện trong ngày
            page=1,
        )

        all_patients = all_res.get("patients", [])
        total_patients = len(all_patients)

        # 5. Phân bổ bệnh nhân vào từng phòng và khoa tương ứng
        patients_by_room_id = {}
        unassigned_patients = []

        for p in all_patients:
            phong_kham = p.get("phong_kham", "").strip().lower()
            matched_room = room_map.get(phong_kham)
            if matched_room:
                r_id = matched_room["room_id"]
                if r_id not in patients_by_room_id:
                    patients_by_room_id[r_id] = []
                patients_by_room_id[r_id].append(p)
            else:
                unassigned_patients.append(p)

        # Xây dựng cấu trúc kết quả phân cấp theo Khoa -> Phòng -> Bệnh nhân
        result_departments = []
        for d in departments:
            dept_patients_count = 0
            dept_rooms = []

            for r in d.get("rooms", []):
                r_id = str(r["room_id"])
                p_list = patients_by_room_id.get(r_id, [])
                dept_patients_count += len(p_list)

                dept_rooms.append({
                    "room_id": r_id,
                    "room_name": r["room_name"],
                    "patient_count": len(p_list),
                    "patients": p_list,
                })

            result_departments.append({
                "dept_id": d["dept_id"],
                "dept_name": d["dept_name"],
                "patient_count": dept_patients_count,
                "rooms_count": len(dept_rooms),
                "rooms": dept_rooms,
            })

        return {
            "success": True,
            "hospital_url": base_url,
            "date": from_date.split(" ")[0],
            "total_departments": len(departments),
            "total_rooms": sum(d.get("rooms_count", 0) for d in departments),
            "total_patients": total_patients,
            "departments": result_departments,
            "unassigned_patients_count": len(unassigned_patients),
            "unassigned_patients": unassigned_patients,
        }


async def export_hospital_patients_excel(
    url: str = _DEFAULT_URL,
    username: str = _DEFAULT_USER,
    password: str = _DEFAULT_PASS,
    room_id: str = "0",
    date_option: str | int = None,
    from_date: str = None,
    to_date: str = None,
    hospital_name: str = "bvlaichau",
    output_dir: str = ".",
    fetch_all: bool = True,
    target_date: str = None,
    progress_callback = None,
) -> dict:
    """
    Lấy TỔNG HỢP toàn bộ danh sách bệnh nhân (tự động hợp nhất cả v1 và v2, khử trùng lặp)
    và xuất ra ĐÚNG 1 FILE EXCEL DUY NHẤT theo định dạng:
    [tên bệnh viện][số lượng].xlsx

    Args:
        url:          URL bệnh viện
        username:     Tài khoản đăng nhập
        password:     Mật khẩu
        room_id:      Mã phòng ('0' = tất cả các phòng khám)
        date_option:  Option chọn khoảng thời gian:
                      '1' hoặc 2021: Từ 00:00 01/01/2021 --> nay
                      '2' hoặc 2022: Từ 00:00 01/01/2022 --> nay
                      '3' hoặc 2023: Từ 00:00 01/01/2023 --> nay
                      '4' hoặc 2024: Từ 00:00 01/01/2024 --> nay
                      '5' hoặc 2025: Từ 00:00 01/01/2025 --> nay
                      '6' hoặc 2026: Từ 00:00 01/01/2026 --> nay
                      'today' hoặc None: Hôm nay (hoặc theo target_date)
        from_date:    Thời gian bắt đầu tuỳ chỉnh
        to_date:      Thời gian kết thúc tuỳ chỉnh
        hospital_name:Tên bệnh viện (mặc định 'bvlaichau')
        output_dir:   Thư mục lưu file
        fetch_all:    Tự động duyệt tất cả các trang nếu nhiều hơn 1 trang
        target_date:  Ngày cụ thể (dd/mm/yyyy) nếu không dùng date_option

    Returns:
        dict: {success, file_path, filename, count, headers, date_label, total_records}
    """
    login_result = await login(url, username, password)
    if not login_result.get("success"):
        return {"success": False, "message": login_result.get("message")}

    cookies = login_result.get("cookies", {})
    location = login_result.get("location", "")
    base_url = login_result.get("base_url") or url.rstrip("/")

    async with httpx.AsyncClient(
        cookies=cookies,
        verify=False,
        timeout=120.0,
        limits=httpx.Limits(max_connections=50, max_keepalive_connections=20),
    ) as client:
        uuid = await _extract_uuid(client, base_url, location)
        if not uuid:
            return {"success": False, "message": "Không tìm thấy token uuid"}

        # 1. Trích xuất tất cả các Stored Procedure từ GridUtil.loadGridBySqlPage trong các trang menu tiếp nhận
        menu_res = await get_reception_menu(client=client, base_url=base_url, load_grids=True)
        if not menu_res.get("has_access") or not menu_res.get("flat_menu"):
            err_msg = f"Tài khoản '{username}' không có quyền truy cập vào phân hệ tiếp nhận của bệnh viện ({url})."
            log.warning(err_msg)
            return {
                "success": False,
                "has_access": False,
                "message": err_msg,
                "count": 0,
                "file_path": None,
                "filename": None,
                "headers": DEFAULT_FIELDS,
                "total_records": 0,
                "procedures_matched": [],
            }

        candidate_procs = []
        for item in menu_res.get("flat_menu", []):
            for g in item.get("grids", []):
                sp = g.get("sql_proc", "").strip()
                if sp and sp not in candidate_procs:
                    candidate_procs.append(sp)

        # Đảm bảo có các procedure tiếp nhận ngoại trú cốt lõi khi tài khoản có quyền tiếp nhận
        for def_p in ["NGT.001.BVTM.01", "NGT.002"]:
            if def_p not in candidate_procs:
                candidate_procs.append(def_p)

        # 2. Bước 2a: Thử nhanh trang 1 (rows=10, fetch_all=False) đồng thời để sàng lọc procedure có dữ liệu bệnh nhân
        async def _probe_proc(p_name: str):
            try:
                res = await get_patients_by_room(
                    client=client,
                    base_url=base_url,
                    uuid=uuid,
                    room_id=room_id,
                    procedure_name=p_name,
                    date_option=date_option,
                    from_date=from_date,
                    to_date=to_date,
                    rows=10,
                    fetch_all=False,
                )
                if not res or not res.get("patients"):
                    return p_name, False, 0, None
                first_raw = res["patients"][0].get("raw", {})
                is_valid = has_required_patient_fields(first_raw)
                return p_name, is_valid, res.get("records", 0), res
            except Exception as e:
                log.debug(f"Probe procedure {p_name} error: {e}")
                return p_name, False, 0, None

        probe_tasks = [_probe_proc(p) for p in candidate_procs]
        probe_results = await asyncio.gather(*probe_tasks)

        valid_procs = []
        for p_name, is_valid, rec_count, p_res in probe_results:
            if is_valid and rec_count > 0:
                valid_procs.append((p_name, rec_count, p_res))

        # Khử trùng lặp giữa NGT.001.BVTM.01 và NGT.002 nếu cùng số lượng bản ghi (v1 vs v2)
        procs_to_fetch = []
        ngt_bvtm = next((p for p in valid_procs if p[0] == "NGT.001.BVTM.01"), None)
        ngt_002 = next((p for p in valid_procs if p[0] == "NGT.002"), None)

        if ngt_bvtm and ngt_002 and ngt_bvtm[1] == ngt_002[1]:
            # Hai procedure trùng khớp dữ liệu, ưu tiên lấy v2 NGT.001.BVTM.01
            procs_to_fetch.append(ngt_bvtm)
            for p in valid_procs:
                if p[0] not in ("NGT.001.BVTM.01", "NGT.002"):
                    procs_to_fetch.append(p)
        else:
            procs_to_fetch = valid_procs

        # 3. Bước 2b: Tải toàn bộ các trang đa luồng song song cho các procedure hợp lệ
        async def _fetch_full_proc(p_tuple):
            p_name, rec_count, initial_res = p_tuple
            # Nếu tổng số bản ghi <= 10 thì đã có đầy đủ từ bước probe
            if rec_count <= 10 and initial_res:
                return p_name, initial_res
            try:
                if progress_callback:
                    try:
                        await progress_callback(f"📥 Đang tải {rec_count:,} bản ghi từ procedure <code>{p_name}</code>...")
                    except Exception:
                        pass
                res = await get_patients_by_room(
                    client=client,
                    base_url=base_url,
                    uuid=uuid,
                    room_id=room_id,
                    procedure_name=p_name,
                    date_option=date_option,
                    from_date=from_date,
                    to_date=to_date,
                    rows=5000,
                    fetch_all=fetch_all,
                    progress_callback=progress_callback,
                )
                return p_name, res
            except Exception as e:
                log.warning(f"Lỗi tải đầy đủ procedure {p_name}: {e}")
                return p_name, None

        full_tasks = [_fetch_full_proc(p) for p in procs_to_fetch]
        full_results = await asyncio.gather(*full_tasks)

        matching_procs_data = []
        valid_proc_names = []

        for p_name, res in full_results:
            if not res:
                continue
            patients_list = res.get("patients", [])
            if patients_list:
                log.info(f"Procedure {p_name} khớp trường bệnh nhân: {len(patients_list)} bản ghi")
                valid_proc_names.append(p_name)
                matching_procs_data.append([p["raw"] for p in patients_list])

        # 4. Hợp nhất và khử trùng lặp toàn bộ bệnh nhân từ tất cả các procedure hợp lệ
        all_merged_raw = []
        if matching_procs_data:
            all_merged_raw = matching_procs_data[0]
            for next_list in matching_procs_data[1:]:
                all_merged_raw = _merge_patient_records(all_merged_raw, next_list)

        if progress_callback:
            try:
                await progress_callback(f"📊 Đã tải xong {len(all_merged_raw):,} bản ghi! Đang tạo file Excel chuẩn...")
            except Exception:
                pass

        excel_res = await asyncio.to_thread(
            export_patients_to_excel,
            all_merged_raw,
            hospital_name=hospital_name,
            output_dir=output_dir,
        )
        excel_res["total_records"] = len(all_merged_raw)
        excel_res["procedures_matched"] = valid_proc_names
        return excel_res


if __name__ == "__main__":
    async def main():
        print("Đang quét bệnh nhân BV Lai Châu...")
        data = await get_hospital_patients_tree()

        print("\n" + "=" * 70)
        print(f"KẾT QUẢ BỆNH NHÂN BV LAI CHÂU - NGÀY {data.get('date')}")
        print(f"Tổng số khoa:      {data.get('total_departments')}")
        print(f"Tổng số phòng:     {data.get('total_rooms')}")
        print(f"Tổng số bệnh nhân: {data.get('total_patients')}")
        print("=" * 70)

        for d in data.get("departments", []):
            if d["patient_count"] > 0:
                print(f"\n🏢 [{d['dept_id']}] {d['dept_name']} -> {d['patient_count']} bệnh nhân:")
                for r in d["rooms"]:
                    if r["patient_count"] > 0:
                        print(f"   🏠 [{r['room_id']}] {r['room_name']}: {r['patient_count']} bệnh nhân")
                        for p in r["patients"]:
                            print(f"       * STT: {p['stt']} | {p['ma_bn']} | {p['ten_bn']} ({p['gioi_tinh']}, {p['ngay_sinh']}) | {p['cong_kham']}")

    asyncio.run(main())
