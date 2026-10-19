import httpx
from urllib.parse import urljoin


# Headers giả lập browser Chrome
_BROWSER_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "vi;q=0.5",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/155.0.0.0 Safari/537.36"
    ),
}

_LOGIN_ENDPOINT = "/servlet/login.ValidateUser"


async def login(url: str, username: str, password: str) -> dict:
    """
    Đăng nhập vào hệ thống VnCare.

    Args:
        url:      URL gốc của bệnh viện, ví dụ: http://bvlamdong.vncare.vn/vnpthis
        username: Tên đăng nhập (txtName)
        password: Mật khẩu (txtPass)

    Returns:
        dict gồm:
            - success    (bool): Đăng nhập thành công hay không
            - message    (str):  Thông báo kết quả
            - cookies    (dict): Cookies phiên làm việc (nếu thành công)
            - status_code (int): HTTP status code
            - location   (str):  Location header từ redirect
    """
    base_url = url.rstrip("/")
    for suffix in [
        "/servlet/login.validateuser",
        "/servlet/login.ValidateUser",
        "/login/login.jsp",
        "/login.jsp",
        "/main/manager.jsp",
    ]:
        if base_url.lower().endswith(suffix.lower()):
            base_url = base_url[: -len(suffix)].rstrip("/")
            break

    try:
        async with httpx.AsyncClient(
            follow_redirects=False,
            timeout=30.0,
            verify=False,
            http1=True,
            http2=False,
        ) as client:

            # --- Bước 1: GET trang chủ để lấy JSESSIONID / SESSIONID và xác định real base URL ---
            resp_home = await client.get(
                base_url + "/",
                headers=_BROWSER_HEADERS,
                follow_redirects=True
            )
            # Server có thể redirect HTTP → HTTPS, lấy URL thực tế
            real_base = str(resp_home.url).rstrip("/")
            login_url = real_base + _LOGIN_ENDPOINT

            # --- Bước 2: POST đăng nhập (không follow redirect để bắt 302) ---
            resp = await client.post(
                login_url,
                data={
                    "srcwidth": "1024",
                    "txtName": username,
                    "txtPass": password,
                },
                headers={
                    **_BROWSER_HEADERS,
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Origin": real_base,
                    "Referer": real_base + "/",
                },
                follow_redirects=False,
            )

            cookies = dict(client.cookies)
            status_code = resp.status_code
            location = resp.headers.get("location", "")

            # --- Bước 3: Kiểm tra kết quả ---
            # Thành công và thất bại đều trả 302, phân biệt qua Location header
            if status_code == 302:
                # SAI MẬT KHẨU: redirect về login.jsp?invalid=1
                if "invalid=1" in location or "login.jsp" in location:
                    return {
                        "success": False,
                        "message": "Sai tài khoản hoặc mật khẩu",
                        "status_code": status_code,
                        "location": location,
                        "cookies": {},
                        "base_url": real_base,
                    }
                # ĐÚNG MẬT KHẨU: redirect về main/manager.jsp hoặc trang chính
                if "main/" in location or "manager.jsp" in location:
                    return {
                        "success": True,
                        "message": "Đăng nhập thành công",
                        "status_code": status_code,
                        "location": location,
                        "cookies": cookies,
                        "base_url": real_base,
                    }

            return {
                "success": False,
                "message": f"Phản hồi không mong đợi (HTTP {status_code})",
                "status_code": status_code,
                "location": location,
                "cookies": {},
            }

    except httpx.ConnectError:
        return {
            "success": False,
            "message": "Không kết nối được tới server",
            "status_code": 0,
            "location": "",
            "cookies": {},
        }
    except httpx.TimeoutException:
        return {
            "success": False,
            "message": "Timeout khi kết nối",
            "status_code": 0,
            "location": "",
            "cookies": {},
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"Lỗi: {e}",
            "status_code": 0,
            "location": "",
            "cookies": {},
        }
