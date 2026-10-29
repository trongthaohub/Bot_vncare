"""Đọc danh sách tỉnh/thành và ánh xạ sang bệnh viện VnCare tương ứng."""

import json
import pathlib
import re
import unicodedata

_ROOT = pathlib.Path(__file__).resolve().parent.parent
_PROVINCES_PATH = _ROOT / "tinh_tp.json"
_HOSPITALS_PATH = _ROOT / "hospitals.json"

# Luôn quét thêm 2 hệ thống dùng chung (chứa tài khoản của nhiều bệnh viện/tỉnh).
SHARED_HOSPITAL_DOMAINS = ("vnpt-his.vncare.vn", "bvdakhoa.vncare.vn")

# Tỉnh có tên không khớp tên bệnh viện trong hospitals.json.
PROVINCE_ALIASES = {
    "thua thien hue": ("hue",),
}

# --- Tỉnh CŨ → tỉnh MỚI (sáp nhập 01/07/2025) ---
# tinh_tp.json là danh sách 63 tỉnh trước sáp nhập, nhưng dữ liệu bệnh nhân
# trên VnCare đã ghi DIACHI theo tên tỉnh mới (ví dụ bệnh nhân Ninh Thuận
# thành "Tỉnh Khánh Hòa", Đắk Nông thành "Tỉnh Lâm Đồng"). Bảng này để vẫn
# lọc được khi người dùng chọn tỉnh cũ: địa chỉ khớp tên cũ (dữ liệu chưa
# chuyển) hoặc tên tỉnh mới chứa tỉnh đó (dữ liệu đã chuyển).
MERGED_INTO = {
    # Miền Bắc
    "ha giang": "Tuyên Quang",
    "tuyen quang": "Tuyên Quang",
    "yen bai": "Lào Cai",
    "lao cai": "Lào Cai",
    "bac kan": "Thái Nguyên",
    "thai nguyen": "Thái Nguyên",
    "vinh phuc": "Phú Thọ",
    "hoa binh": "Phú Thọ",
    "phu tho": "Phú Thọ",
    "bac giang": "Bắc Ninh",
    "bac ninh": "Bắc Ninh",
    "thai binh": "Hưng Yên",
    "hung yen": "Hưng Yên",
    "hai duong": "Hải Phòng",
    "hai phong": "Hải Phòng",
    "ha nam": "Ninh Bình",
    "nam dinh": "Ninh Bình",
    "ninh binh": "Ninh Bình",
    # Miền Trung
    "quang binh": "Quảng Trị",
    "quang tri": "Quảng Trị",
    "quang nam": "Đà Nẵng",
    "da nang": "Đà Nẵng",
    "kon tum": "Quảng Ngãi",
    "quang ngai": "Quảng Ngãi",
    "binh dinh": "Gia Lai",
    "gia lai": "Gia Lai",
    "ninh thuan": "Khánh Hòa",
    "khanh hoa": "Khánh Hòa",
    "dak nong": "Lâm Đồng",
    "binh thuan": "Lâm Đồng",
    "lam dong": "Lâm Đồng",
    "phu yen": "Đắk Lắk",
    "dak lak": "Đắk Lắk",
    # Miền Nam
    "binh duong": "Hồ Chí Minh",
    "ba ria - vung tau": "Hồ Chí Minh",
    "ho chi minh": "Hồ Chí Minh",
    "binh phuoc": "Đồng Nai",
    "dong nai": "Đồng Nai",
    "long an": "Tây Ninh",
    "tay ninh": "Tây Ninh",
    "soc trang": "Cần Thơ",
    "hau giang": "Cần Thơ",
    "can tho": "Cần Thơ",
    "ben tre": "Vĩnh Long",
    "tra vinh": "Vĩnh Long",
    "vinh long": "Vĩnh Long",
    "tien giang": "Đồng Tháp",
    "dong thap": "Đồng Tháp",
    "kien giang": "An Giang",
    "an giang": "An Giang",
    "bac lieu": "Cà Mau",
    "ca mau": "Cà Mau",
    # Thành phố trực thuộc TW
    "thua thien hue": "Huế",
    "hue": "Huế",
    "ha noi": "Hà Nội",
}

_SPECIAL_CHARS = str.maketrans({"đ": "d", "Đ": "D"})


def normalize_text(value: str) -> str:
    """Bỏ dấu tiếng Việt, lowercase, gộp khoảng trắng."""
    if not value:
        return ""
    text = str(value).translate(_SPECIAL_CHARS)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.lower().split())


def _compact(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))


def load_provinces() -> list[dict]:
    """Trả về danh sách tỉnh/thành, sắp xếp theo tên."""
    with _PROVINCES_PATH.open(encoding="utf-8") as f:
        data = json.load(f)
    provinces = [dict(info, code=str(code)) for code, info in data.items()]
    provinces.sort(key=lambda p: normalize_text(p.get("name", "")))
    return provinces


def find_province(code: str) -> dict | None:
    for p in load_provinces():
        if p.get("code") == str(code):
            return p
    return None


def load_hospitals() -> list[dict]:
    with _HOSPITALS_PATH.open(encoding="utf-8") as f:
        return json.load(f).get("hospitals", [])


def _hospital_core(name: str) -> str:
    core = _compact(name)
    for prefix in ("staging", "benhvien", "bv"):
        if core.startswith(prefix) and len(core) > len(prefix):
            core = core[len(prefix):]
    return core


def province_hospitals(province: dict) -> list[dict]:
    """Trả về bệnh viện của tỉnh + các hệ thống dùng chung (VNPT-HIS, BVDAKHOA)."""
    hospitals = load_hospitals()
    province_name = normalize_text(province.get("name", ""))
    targets = {
        _compact(province.get("name", "")),
        _compact(province.get("name_with_type", "")),
        _compact(province.get("slug", "")),
    }
    for prefix in ("tinh", "thanhpho"):
        targets.discard(prefix)
    for alias in PROVINCE_ALIASES.get(province_name, ()):
        targets.add(_compact(alias))
    targets.discard("")

    matched: list[dict] = []
    seen: set[str] = set()

    def _add(h: dict) -> None:
        domain = (h.get("domain") or "").lower()
        if domain and domain not in seen:
            seen.add(domain)
            matched.append(h)

    for h in hospitals:
        core = _hospital_core(h.get("name", ""))
        if not core:
            continue
        for t in targets:
            if core == t or (len(t) >= 5 and t in core) or (len(core) >= 5 and core in t):
                _add(h)
                break

    for h in hospitals:
        if (h.get("domain") or "").lower() in SHARED_HOSPITAL_DOMAINS:
            _add(h)

    return matched


def province_strict_names(province: dict) -> set[str]:
    """Tên tỉnh đúng như người dùng đã chọn (chưa tính tỉnh mới sau sáp nhập)."""
    name = normalize_text(province.get("name", ""))
    names = {name, normalize_text(province.get("name_with_type", ""))}
    for n in list(names):
        for prefix in ("tinh ", "thanh pho ", "tp. ", "tp "):
            if n.startswith(prefix):
                names.add(n[len(prefix):])
    names.update(normalize_text(a) for a in PROVINCE_ALIASES.get(name, ()))
    names.discard("")
    return names


def province_search_names(province: dict) -> set[str]:
    """Các tên tỉnh có thể xuất hiện trong DIACHI (tên cũ, tên mới sau sáp nhập)."""
    names = province_strict_names(province)
    merged = MERGED_INTO.get(normalize_text(province.get("name", "")))
    if merged:
        names.add(normalize_text(merged))
    return names


def _address_has(addr: str, candidate: str) -> bool:
    pattern = rf"(?<![a-z0-9]){re.escape(candidate)}(?![a-z0-9])"
    return bool(re.search(pattern, addr))


def address_match_level(address: str, province: dict) -> str | None:
    """"strict" | "merged" | None.

    - strict: DIACHI ghi đúng tên tỉnh người dùng chọn.
    - merged: DIACHI chỉ ghi tên tỉnh MỚI chứa tỉnh đó (dữ liệu đã sáp nhập
      01/07/2025 nên không còn phân biệt được ở mức tỉnh).
    """
    addr = normalize_text(address)
    if not addr:
        return None
    strict = province_strict_names(province)
    for candidate in strict:
        if _address_has(addr, candidate):
            return "strict"
    for candidate in province_search_names(province) - strict:
        if _address_has(addr, candidate):
            return "merged"
    return None


def address_matches_province(address: str, province: dict) -> bool:
    """True nếu chuỗi địa chỉ (DIACHI) chứa tên tỉnh/thành đã chọn.

    Chấp nhận cả tên tỉnh cũ (dữ liệu chưa chuyển) và tên tỉnh mới sau sáp
    nhập 01/07/2025 (xem MERGED_INTO).
    """
    return address_match_level(address, province) is not None
