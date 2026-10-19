import asyncio
import json
import logging
import re
from urllib.parse import parse_qs, urljoin, urlparse
import httpx

log = logging.getLogger(__name__)

_DEFAULT_URL = "https://bvlaichau.vncare.vn/vnpthis"
_DEFAULT_MENU_PATH = "common/script/menu/1_2_vn.js"

_HEADERS = {
    "Accept": "*/*",
    "Accept-Language": "vi;q=0.5",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/155.0.0.0 Safari/537.36"
    ),
    "X-Requested-With": "XMLHttpRequest",
}


def _flatten_menu_tree(nodes: list, parent_path: str = "") -> list:
    """Duyệt đệ quy cây menu để tạo danh sách phẳng các chức năng có liên kết."""
    flat = []
    for node in nodes:
        title = node.get("text", "").strip()
        curr_path = f"{parent_path} > {title}" if parent_path else title
        hlink = node.get("hlink", "").strip()
        children = node.get("children", [])

        func_param = ""
        if hlink and hlink != "#":
            parsed = urlparse(hlink)
            qs = parse_qs(parsed.query)
            if "func" in qs:
                func_param = qs["func"][0]

            flat.append({
                "id": node.get("id", ""),
                "title": title,
                "path": curr_path,
                "hlink": hlink,
                "func": func_param,
                "icon": node.get("icon", ""),
                "options": node.get("options", ""),
            })

        if children:
            flat.extend(_flatten_menu_tree(children, curr_path))

    return flat


def _filter_reception_nodes(menu_tree: list) -> list:
    """
    Lọc trích xuất các nhánh menu thuộc phân hệ 'Tiếp nhận'.
    Bao gồm Tiếp nhận khám bệnh (nhóm 201) và các màn hình tiếp nhận ngoại trú.
    """
    reception_nodes = []
    seen_ids = set()

    for root in menu_tree:
        root_title = root.get("text", "").lower()

        # Nhánh Ngoại trú -> Tiếp nhận khám bệnh
        if "ngoại trú" in root_title:
            for sub in root.get("children", []):
                sub_title = sub.get("text", "").lower()
                if "tiếp nhận" in sub_title and sub.get("id") not in seen_ids:
                    seen_ids.add(sub.get("id"))
                    reception_nodes.append(sub)

        # Các nhánh gốc khác nếu có Tiếp nhận
        elif "tiếp nhận" in root_title and root.get("id") not in seen_ids:
            seen_ids.add(root.get("id"))
            reception_nodes.append(root)

    return reception_nodes


def extract_grids_from_js(js_code: str) -> list[dict]:
    """
    Phân tích mã nguồn file Javascript để bóc tách các lệnh GridUtil.loadGridBySqlPage.

    Cú pháp điển hình trong VnCare:
        GridUtil.loadGridBySqlPage("grdDanhSachTiepNhan", "NGT.001.BVTM.01", sql_par)
        GridUtil.loadGridBySqlPage(_gridId, "NGT02K041.DSSLPK", sql_par)

    Returns:
        list[dict]: Danh sách các grid gồm: {grid_id, sql_proc, params_var, raw_call}
    """
    results = []
    # Regex bắt tất cả các lệnh gọi GridUtil.loadGridBySqlPage(...)
    pattern = re.compile(
        r"GridUtil\.loadGridBySqlPage\s*\(\s*([^,\)]+)\s*,\s*([^,\)]+)(?:\s*,\s*([^,\)]+))?",
        re.MULTILINE,
    )

    for match in pattern.finditer(js_code):
        raw_grid = match.group(1).strip().strip("\"'")
        raw_sql = match.group(2).strip().strip("\"'")
        raw_params = match.group(3).strip() if match.group(3) else ""

        results.append({
            "grid_id": raw_grid,
            "sql_proc": raw_sql,
            "params_var": raw_params,
            "raw_call": match.group(0).strip(),
        })

    return results


async def get_page_grid_info(
    client: httpx.AsyncClient,
    base_url: str,
    hlink_or_func: str,
) -> dict:
    """
    Lấy thông tin GridUtil.loadGridBySqlPage của một trang thông qua URL menu + .js.

    Args:
        client:        httpx.AsyncClient đang giữ phiên đăng nhập
        base_url:      Base URL bệnh viện (vd: https://bvlaichau.vncare.vn/vnpthis)
        hlink_or_func: Link menu (vd: /vnpthis/main/manager.jsp?func=../ngoaitru/NGT01T001_tiepnhan_ngt_vs2)
                       hoặc mã hàm (vd: ../ngoaitru/NGT01T001_tiepnhan_ngt_vs2)

    Returns:
        dict: {js_url, status_code, grids, total_grids}
    """
    func_path = hlink_or_func
    if "func=" in hlink_or_func:
        parsed = urlparse(hlink_or_func)
        qs = parse_qs(parsed.query)
        func_path = qs.get("func", [""])[0]

    # Chuẩn hóa đường dẫn file .js từ func
    clean_func = func_path.replace("../", "").lstrip("/")
    if not clean_func.endswith(".js"):
        clean_func = f"{clean_func}.js"

    root_url = base_url.split("/common/")[0].rstrip("/")
    js_url = f"{root_url}/{clean_func}"

    try:
        resp = await client.get(js_url, headers=_HEADERS)
        if resp.status_code == 200:
            grids = extract_grids_from_js(resp.text)
            return {
                "js_url": js_url,
                "status_code": 200,
                "success": True,
                "total_grids": len(grids),
                "grids": grids,
            }
        else:
            return {
                "js_url": js_url,
                "status_code": resp.status_code,
                "success": False,
                "total_grids": 0,
                "grids": [],
            }
    except Exception as e:
        return {
            "js_url": js_url,
            "status_code": 0,
            "success": False,
            "error": str(e),
            "total_grids": 0,
            "grids": [],
        }


async def check_reception_access(
    client: httpx.AsyncClient,
    base_url: str = _DEFAULT_URL,
) -> bool:
    """
    Kiểm tra tài khoản hiện tại có quyền truy cập vào màn hình tiếp nhận hay không.

    Returns:
        bool: True nếu tài khoản có quyền truy cập, False nếu bị chặn/chuyển hướng hoặc không có quyền.
    """
    test_urls = [
        f"{base_url.rstrip('/')}/main/manager.jsp?func=../ngoaitru/NGT01T001_tiepnhan_ngt_vs2",
        f"{base_url.rstrip('/')}/main/manager.jsp?func=../ngoaitru/NGT01T001_tiepnhan_ngt",
    ]
    for test_url in test_urls:
        try:
            resp = await client.get(test_url, headers=_HEADERS, follow_redirects=True)
            if resp.status_code == 200:
                resp_str = str(resp.url).lower()
                if "login" in resp_str or "error" in resp_str:
                    continue
                if "không có quyền" in resp.text.lower() or "access denied" in resp.text.lower():
                    continue
                return True
        except Exception as e:
            log.warning(f"Lỗi kiểm tra quyền tiếp nhận tại {test_url}: {e}")
    return False


async def get_reception_menu(
    cookies: dict | str = None,
    base_url: str = _DEFAULT_URL,
    login_result: dict = None,
    client: httpx.AsyncClient = None,
    load_grids: bool = True,
) -> dict:
    """
    Chỉ trả về menu của phần TIẾP NHẬN NẾU TÀI KHOẢN CÓ QUYỀN TRUY CẬP.
    Đồng thời bóc tách toàn bộ GridUtil.loadGridBySqlPage của từng trang qua file JS tương ứng.

    Args:
        cookies:      Cookies phiên đăng nhập
        base_url:     URL gốc của bệnh viện
        login_result: Kết quả dict từ login()
        client:       httpx.AsyncClient tái sử dụng nếu có
        load_grids:   Nếu True, tự động tải file .js của từng trang để lấy GridUtil.loadGridBySqlPage

    Returns:
        dict: {
            success: bool,
            has_access: bool,
            message: str,
            total_items: int,
            menu: list,       # Cây menu phân cấp
            flat_menu: list   # Danh sách phẳng các trang kèm 'grids'
        }
    """
    return await get_menu(
        cookies=cookies,
        base_url=base_url,
        login_result=login_result,
        client=client,
        only_reception=True,
        load_grids=load_grids,
    )


async def get_menu(
    cookies: dict | str = None,
    base_url: str = _DEFAULT_URL,
    menu_path: str = _DEFAULT_MENU_PATH,
    login_result: dict = None,
    client: httpx.AsyncClient = None,
    only_reception: bool = False,
    load_grids: bool = False,
) -> dict:
    """
    Lấy cây menu chức năng của hệ thống VnCare từ script menu .js hoặc _menuData phiên làm việc.
    Nếu only_reception=True: chỉ trả về menu phần tiếp nhận nếu tài khoản có quyền truy cập.
    Nếu load_grids=True: tự động truy cập vào file .js của từng page để lấy GridUtil.loadGridBySqlPage.

    Args:
        cookies:        Dictionary chứa cookies sau khi đăng nhập
        base_url:       URL gốc của bệnh viện hoặc URL trực tiếp tới file .js
        menu_path:      Đường dẫn tới file menu (mặc định: common/script/menu/1_2_vn.js)
        login_result:   Kết quả dict từ hàm login()
        client:         httpx.AsyncClient tái sử dụng nếu có
        only_reception: Nếu True, chỉ trả về menu phần tiếp nhận khi có quyền truy cập
        load_grids:     Nếu True, bóc tách GridUtil.loadGridBySqlPage của từng page

    Returns:
        dict: {success, has_access, menu, flat_menu, total_items, ...}
    """
    if login_result and isinstance(login_result, dict):
        cookies = login_result.get("cookies", cookies)

    if isinstance(cookies, str):
        cookie_dict = {}
        for part in cookies.split(";"):
            if "=" in part:
                k, v = part.strip().split("=", 1)
                cookie_dict[k] = v
        cookies = cookie_dict

    cookies = cookies or {}

    clean_base = base_url.strip()
    if clean_base.endswith(".js"):
        menu_url = clean_base
        root_url = clean_base.split("/common/")[0]
    else:
        root_url = clean_base.rstrip("/")
        clean_path = menu_path.lstrip("/")
        menu_url = f"{root_url}/{clean_path}"

    async def _fetch(c: httpx.AsyncClient) -> dict:
        # 1. Kiểm tra quyền truy cập phần tiếp nhận
        if only_reception:
            has_perm = await check_reception_access(c, base_url=root_url)
            if not has_perm:
                return {
                    "success": False,
                    "has_access": False,
                    "message": "Tài khoản không có quyền truy cập phân hệ Tiếp nhận.",
                    "total_items": 0,
                    "menu": [],
                    "flat_menu": [],
                }

        # 2. Thử bóc tách cây menu thực tế của người dùng từ manager.jsp trước (_menuData)
        menu_tree = None
        for page_url in [
            f"{root_url}/main/manager.jsp?func=../ngoaitru/NGT01T001_tiepnhan_ngt_vs2",
            f"{root_url}/main/manager.jsp?func=../ngoaitru/NGT01T001_tiepnhan_ngt",
            f"{root_url}/main/manager.jsp?func=../admin/SelDept",
        ]:
            try:
                page_resp = await c.get(page_url, headers=_HEADERS, follow_redirects=True)
                if page_resp.status_code == 200:
                    match = re.search(r"var\s+_menuData\s*=\s*(\[[\s\S]*?\]);", page_resp.text)
                    if match:
                        raw_data = json.loads(match.group(1))
                        active_data = [
                            m for m in raw_data
                            if m.get("text") not in ("Popup", "Dược popup UI")
                        ]
                        if only_reception:
                            rec_nodes = _filter_reception_nodes(active_data)
                            if rec_nodes:
                                menu_tree = active_data
                                break
                        else:
                            menu_tree = active_data
                            break
            except Exception:
                pass

        # 3. Nếu không có _menuData phân quyền động hoặc _menuData chưa có nhánh tiếp nhận, tải menu tĩnh
        if menu_tree is None or (only_reception and not _filter_reception_nodes(menu_tree)):
            resp = await c.get(
                menu_url,
                headers={**_HEADERS, "Referer": f"{root_url}/main/manager.jsp"},
                follow_redirects=True,
            )
            if resp.status_code != 200:
                return {
                    "success": False,
                    "has_access": False,
                    "status_code": resp.status_code,
                    "message": f"Không thể tải menu, mã phản hồi: {resp.status_code}",
                    "menu_url": menu_url,
                    "menu": [],
                    "flat_menu": [],
                }

            text = resp.text.strip()
            if text.startswith("var ") and "=" in text:
                text = text.split("=", 1)[1].strip()
            if text.endswith(";"):
                text = text[:-1].strip()

            try:
                menu_tree = json.loads(text)
            except Exception as e:
                match = re.search(r"(\[[\s\S]*\])", text)
                if match:
                    menu_tree = json.loads(match.group(1))
                else:
                    return {
                        "success": False,
                        "has_access": False,
                        "message": f"Không thể parse JSON menu: {e}",
                        "menu_url": menu_url,
                        "menu": [],
                        "flat_menu": [],
                    }

        # 4. Lọc danh sách menu
        if only_reception:
            target_nodes = _filter_reception_nodes(menu_tree)
            if not target_nodes:
                return {
                    "success": False,
                    "has_access": False,
                    "message": "Tài khoản không có quyền truy cập phân hệ Tiếp nhận.",
                    "total_items": 0,
                    "menu": [],
                    "flat_menu": [],
                }
            module_name = "Tiếp nhận"
            msg = "Tài khoản có quyền truy cập phân hệ Tiếp nhận."
        else:
            target_nodes = menu_tree
            module_name = "Tất cả"
            msg = "Đã lấy danh sách menu thành công."

        flat_list = _flatten_menu_tree(target_nodes)

        # 4. Nếu load_grids=True: truy cập URL menu + .js để lấy GridUtil.loadGridBySqlPage
        if load_grids and flat_list:
            log.info(f"Đang bóc tách GridUtil cho {len(flat_list)} trang...")
            for item in flat_list:
                func = item.get("func", "")
                hlink = item.get("hlink", "")
                if func or ("func=" in hlink):
                    grid_info = await get_page_grid_info(c, root_url, func or hlink)
                    item["js_url"] = grid_info.get("js_url", "")
                    item["grids"] = grid_info.get("grids", [])
                else:
                    item["js_url"] = ""
                    item["grids"] = []

        return {
            "success": True,
            "has_access": True,
            "module": module_name,
            "message": msg,
            "menu_url": menu_url,
            "total_items": len(flat_list),
            "menu": target_nodes,
            "flat_menu": flat_list,
        }

    if client:
        return await _fetch(client)

    async with httpx.AsyncClient(
        cookies=cookies,
        verify=False,
        timeout=45.0,
        http1=True,
    ) as new_client:
        return await _fetch(new_client)


def search_menu(menu_data: dict | list, keyword: str) -> list:
    """Tìm kiếm nhanh các mục menu theo từ khóa tên, mã hàm, hoặc Stored Procedure."""
    if isinstance(menu_data, dict):
        flat_list = menu_data.get("flat_menu", [])
    elif isinstance(menu_data, list):
        if menu_data and "path" in menu_data[0]:
            flat_list = menu_data
        else:
            flat_list = _flatten_menu_tree(menu_data)
    else:
        return []

    kw = keyword.lower().strip()
    matched = []
    for item in flat_list:
        title = item.get("title", "").lower()
        path = item.get("path", "").lower()
        func = item.get("func", "").lower()
        hlink = item.get("hlink", "").lower()

        # Tìm trong các grid sql_proc nếu có
        grids_match = False
        for g in item.get("grids", []):
            if kw in g.get("sql_proc", "").lower() or kw in g.get("grid_id", "").lower():
                grids_match = True
                break

        if kw in title or kw in path or kw in func or kw in hlink or grids_match:
            matched.append(item)

    return matched


if __name__ == "__main__":
    import os
    import sys

    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

    from libs.login import login
    from Controller.bvlaichau import _DEFAULT_URL, _DEFAULT_USER, _DEFAULT_PASS

    async def main():
        print("Đang đăng nhập...")
        log_res = await login(_DEFAULT_URL, _DEFAULT_USER, _DEFAULT_PASS)
        if not log_res.get("success"):
            print("Đăng nhập thất bại!")
            return

        print("Đang kiểm tra quyền và lấy menu Tiếp nhận kèm GridUtil.loadGridBySqlPage...")
        res = await get_reception_menu(
            cookies=log_res["cookies"],
            base_url=f"{_DEFAULT_URL}/common/script/menu/1_2_vn.js",
            load_grids=True,
        )

        print("\n" + "=" * 80)
        print("KẾT QUẢ MENU TIẾP NHẬN KÈM GRID BINDING")
        print(f"Có quyền truy cập: {res.get('has_access')}")
        print(f"Tổng số trang:     {res.get('total_items')}")
        print("=" * 80)

        for idx, item in enumerate(res.get("flat_menu", []), 1):
            print(f"\n{idx:2d}. [{item['id']}] {item['title']}")
            print(f"    Menu URL: {item['hlink']}")
            print(f"    JS File:  {item.get('js_url')}")
            grids = item.get("grids", [])
            if grids:
                print(f"    -> Có {len(grids)} lệnh GridUtil.loadGridBySqlPage:")
                for g in grids:
                    print(f"       * Grid: {g['grid_id']} | SQL Proc: {g['sql_proc']} | Params: {g['params_var']}")
            else:
                print("    -> (Không có GridUtil.loadGridBySqlPage)")

    asyncio.run(main())
