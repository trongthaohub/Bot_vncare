import json
import logging
import re
from urllib.parse import urljoin
import httpx

from database import Database
from libs.login import login

log = logging.getLogger(__name__)

# Headers giả lập Chrome browser
_HEADERS = {
    "Accept": "*/*",
    "Accept-Language": "vi;q=0.5",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/155.0.0.0 Safari/537.36"
    ),
}


def _normalize_base_url(url: str) -> str:
    """Chuẩn hóa URL bệnh viện về dạng base URL (vd: https://bvlamdong.vncare.vn/vnpthis)."""
    base = url.strip().rstrip("/")
    if "/servlet/" in base.lower():
        base = re.split(r"/servlet/", base, flags=re.IGNORECASE)[0]
    elif "/login/" in base.lower():
        base = re.split(r"/login/", base, flags=re.IGNORECASE)[0]
    elif "/main/" in base.lower():
        base = re.split(r"/main/", base, flags=re.IGNORECASE)[0]
    return base.rstrip("/")


async def sel_dept(
    url: str,
    login_result: dict = None,
    username: str = None,
    password: str = None,
    db: Database = None,
    force_refresh: bool = False,
) -> list:
    """
    Lấy danh sách khoa và phòng của bệnh viện.
    Cập nhật vào database nếu chưa có trong database.

    Args:
        url:           URL bệnh viện (vd: https://bvgialai.vncare.vn/vnpthis)
        login_result:  Kết quả trả về từ libs.login.login(...)
        username:      Tên tài khoản (nếu chưa login trước)
        password:      Mật khẩu (nếu chưa login trước)
        db:            Database instance (nếu None sẽ tự khởi tạo)
        force_refresh: Buộc quét lại từ API kể cả khi DB đã có dữ liệu

    Returns:
        dict:
            - success (bool): Thành công hay thất bại
            - dept_count (int): Số lượng khoa
            - room_count (int): Tổng số lượng phòng
            - total_departments (int): Số lượng khoa (alias)
            - total_rooms (int): Tổng số lượng phòng (alias)
            - departments (list): Danh sách các khoa, mỗi khoa kèm danh sách phòng
    """
    if db is None:
        db = Database()

    # Nếu chưa có login_result nhưng có username & password -> thực hiện login
    if not login_result and username and password:
        login_result = await login(url=url, username=username, password=password)

    if not login_result or not login_result.get("success", False):
        msg = login_result.get("message", "Login thất bại") if login_result else "Thiếu kết quả đăng nhập"
        log.warning(f"sel_dept: {msg}")
        return {
            "success": False,
            "message": msg,
            "dept_count": 0,
            "room_count": 0,
            "total_departments": 0,
            "total_rooms": 0,
            "departments": [],
        }

    cookies = login_result.get("cookies", {})
    location = login_result.get("location", "")

    if not cookies:
        log.warning("sel_dept: Cookies phiên làm việc trống.")
        return {
            "success": False,
            "message": "Cookies phiên làm việc trống",
            "dept_count": 0,
            "room_count": 0,
            "total_departments": 0,
            "total_rooms": 0,
            "departments": [],
        }

    base_url = _normalize_base_url(url)

    async with httpx.AsyncClient(
        cookies=cookies,
        timeout=30.0,
        verify=False,
        http1=True,
        http2=False,
    ) as client:
        try:
            # 1. Tải trang quản trị (manager.jsp) để lấy context (hospital_id, user_id, uuid)
            manager_url = urljoin(base_url, location) if location else base_url + "/main/manager.jsp"
            resp_mgr = await client.get(manager_url, headers=_HEADERS, follow_redirects=True)
            html = resp_mgr.text

            uuid_match = re.search(r"var\s+uuid\s*=\s*'([^']+)'", html)
            hosp_match = re.search(r"var\s+hospital_id\s*=\s*'([^']+)'", html)
            user_match = re.search(r"var\s+user_id\s*=\s*'([^']+)'", html)

            uuid = uuid_match.group(1) if uuid_match else ""
            hospital_id = hosp_match.group(1) if hosp_match else ""
            user_id = user_match.group(1) if user_match else ""

            if not hospital_id or not uuid or not user_id:
                log.error("sel_dept: Không tìm thấy hospital_id/user_id/uuid trong trang manager.")
                return {
                    "success": False,
                    "message": "Không tìm thấy token phiên trong trang manager",
                    "dept_count": 0,
                    "room_count": 0,
                    "total_departments": 0,
                    "total_rooms": 0,
                    "departments": [],
                }

            # 2. Kiểm tra nếu đã có trong database và không yêu cầu force_refresh -> trả về từ database
            if not force_refresh and db.has_hospital_departments(hospital_id):
                log.info(f"sel_dept: Bệnh viện (hospital_id={hospital_id}) đã có trong DB. Trả về từ DB.")
                depts_from_db = db.get_full_departments_tree(hospital_id)
                total_rooms = sum(d.get("rooms_count", len(d.get("rooms", []))) for d in depts_from_db)
                return {
                    "success": True,
                    "hospital_id": hospital_id,
                    "dept_count": len(depts_from_db),
                    "room_count": total_rooms,
                    "total_departments": len(depts_from_db),
                    "total_rooms": total_rooms,
                    "departments": depts_from_db,
                }

            # 3. Chưa có trong database -> Quét từ API VnCare
            log.info(f"sel_dept: Chưa có trong DB (hoặc force_refresh), bắt đầu quét API cho hospital_id={hospital_id}...")
            real_base = str(resp_mgr.url).split("/main/")[0].rstrip("/")
            rest_url = real_base + "/RestService"

            # Lấy danh sách Khoa
            payload_khoa = {
                "func": "dbCALL_SP_R",
                "params": ["", "DEPT.P01", f"4${user_id}${hospital_id}", 0],
                "uuid": uuid,
            }

            resp_dept = await client.post(
                rest_url,
                content=json.dumps(payload_khoa),
                headers={
                    **_HEADERS,
                    "Content-Type": "text/plain; charset=utf-8",
                    "Referer": str(resp_mgr.url),
                },
            )

            try:
                dept_data = resp_dept.json()
            except Exception:
                dept_data = json.loads(resp_dept.text)

            result_raw = dept_data.get("result", "[]")
            departments_raw = json.loads(result_raw) if isinstance(result_raw, str) else result_raw

            departments_list = []

            # Lấy danh sách Phòng của từng Khoa
            for dept in departments_raw:
                if isinstance(dept, list):
                    dept_id = str(dept[0])
                    dept_name = str(dept[1])
                else:
                    dept_id = str(dept.get("ID") or dept.get("id") or dept.get("VALUE") or dept.get("value"))
                    dept_name = str(dept.get("NAME") or dept.get("name") or dept.get("TEXT") or dept.get("text"))

                payload_phong = {
                    "func": "dbCALL_SP_R",
                    "params": ["", "DEPT.P01", f"5${user_id}${dept_id}", 0],
                    "uuid": uuid,
                }

                resp_room = await client.post(
                    rest_url,
                    content=json.dumps(payload_phong),
                    headers={
                        **_HEADERS,
                        "Content-Type": "text/plain; charset=utf-8",
                        "Referer": str(resp_mgr.url),
                    },
                )

                try:
                    room_data = resp_room.json()
                except Exception:
                    room_data = json.loads(resp_room.text)

                room_raw = room_data.get("result", "[]")
                rooms_raw = json.loads(room_raw) if isinstance(room_raw, str) else room_raw

                room_list = []
                for room in rooms_raw:
                    if isinstance(room, list):
                        room_id = str(room[0])
                        room_name = str(room[1])
                    else:
                        room_id = str(room.get("ID") or room.get("id") or room.get("VALUE") or room.get("value"))
                        room_name = str(room.get("NAME") or room.get("name") or room.get("TEXT") or room.get("text"))

                    room_list.append({
                        "room_id": room_id,
                        "room_name": room_name,
                    })

                departments_list.append({
                    "dept_id": dept_id,
                    "dept_name": dept_name,
                    "rooms_count": len(room_list),
                    "rooms": room_list,
                })

            # 4. Cập nhật vào database
            saved_rooms = db.save_departments_and_rooms(
                hospital_id=hospital_id,
                hospital_url=real_base,
                departments=departments_list,
            )
            log.info(f"sel_dept: Đã cập nhật vào DB {len(departments_list)} khoa và {saved_rooms} phòng (hospital_id={hospital_id}).")

            total_rooms = sum(d["rooms_count"] for d in departments_list)
            return {
                "success": True,
                "hospital_id": hospital_id,
                "dept_count": len(departments_list),
                "room_count": total_rooms,
                "total_departments": len(departments_list),
                "total_rooms": total_rooms,
                "departments": departments_list,
            }

        except Exception as e:
            log.exception(f"sel_dept: Lỗi khi xử lý khoa/phòng: {e}")
            return {
                "success": False,
                "message": f"Lỗi xử lý khoa/phòng: {e}",
                "dept_count": 0,
                "room_count": 0,
                "total_departments": 0,
                "total_rooms": 0,
                "departments": [],
            }


# Alias tương thích
get_departments_and_rooms = sel_dept


if __name__ == "__main__":
    import asyncio
    import sys

    h_url = sys.argv[1] if len(sys.argv) > 1 else "https://bvgialai.vncare.vn/vnpthis"
    h_user = sys.argv[2] if len(sys.argv) > 2 else "bvdk.lthuong"
    h_pass = sys.argv[3] if len(sys.argv) > 3 else "Huong1980@"

    async def main():
        print("1. Đăng nhập qua libs.login...")
        login_res = await login(h_url, h_user, h_pass)

        print("\n2. Gọi sel_dept(url, login_result)...")
        depts = await sel_dept(url=h_url, login_result=login_res)
        print("Khoa và phòng trả về:")
        print(json.dumps(depts, ensure_ascii=False, indent=2))

    asyncio.run(main())
