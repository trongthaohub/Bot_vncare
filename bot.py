import asyncio
from collections import Counter
import io
import json
import logging
import logging.handlers
import pathlib
import time
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
import os
from telegram.request import HTTPXRequest
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from database import Database
from libs.login import login
from libs.sel_dept import sel_dept
from Controller.bvlaichau import export_hospital_patients_excel
import config

# Setup logging with colored console output and plain file output
logdir_path = pathlib.Path(__file__).parent.joinpath("logs").absolute()
logfile_path = logdir_path.joinpath("bot.log")
if not logdir_path.exists():
    logdir_path.mkdir(parents=True, exist_ok=True)

loglevels = {"debug": logging.DEBUG, "error": logging.ERROR, "fatal": logging.FATAL, "info": logging.INFO}
loglevel = loglevels.get(config.LOGLEVEL.lower(), logging.INFO)

logger = logging.getLogger()
logger.setLevel(loglevel)

# File Handler (Không chứa mã màu ANSI để file log sạch sẽ)
logfile_handler = logging.handlers.WatchedFileHandler(logfile_path, "a", "utf-8")
file_formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logfile_handler.setFormatter(file_formatter)
logger.addHandler(logfile_handler)

# Custom Colored Formatter cho Console (StreamHandler)
class ColoredFormatter(logging.Formatter):
    grey = "\x1b[38;20m"
    green = "\x1b[32;20m"
    yellow = "\x1b[33;20m"
    red = "\x1b[31;20m"
    bold_red = "\x1b[31;1m"
    cyan = "\x1b[36;20m"
    reset = "\x1b[0m"

    FORMATS = {
        logging.DEBUG: cyan + "%(asctime)s - %(levelname)s - %(message)s" + reset,
        logging.INFO: green + "%(asctime)s - %(levelname)s - %(message)s" + reset,
        logging.WARNING: yellow + "%(asctime)s - %(levelname)s - %(message)s" + reset,
        logging.ERROR: red + "%(asctime)s - %(levelname)s - %(message)s" + reset,
        logging.CRITICAL: bold_red + "%(asctime)s - %(levelname)s - %(message)s" + reset,
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno, self.cyan)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)

stream_handler = logging.StreamHandler()
stream_handler.setFormatter(ColoredFormatter())
logger.addHandler(stream_handler)

log = logging.getLogger(__name__)

# Khởi tạo Database instance
db = Database()
START_TIME = time.time()

def format_uptime(seconds: float) -> str:
    seconds = int(seconds)
    days, remainder = divmod(seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    parts = []
    if days > 0:
        parts.append(f"{days} ngày")
    if hours > 0:
        parts.append(f"{hours} giờ")
    if minutes > 0:
        parts.append(f"{minutes} phút")
    parts.append(f"{seconds} giây")
    return " ".join(parts)

def is_allowed_chat(update: Update) -> bool:
    chat = update.effective_chat
    if chat and chat.type in ["group", "supergroup"]:
        if config.ALLOWED_GROUPS and chat.id not in config.ALLOWED_GROUPS:
            return False
    return True

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update):
        return
    user = update.effective_user
    db.save_user(user.id, user.username, user.first_name)
    
    # Ưu tiên hiển thị @username, nếu không có thì lấy first_name
    name_display = f"@{user.username}" if user.username else user.first_name
    
    uptime_str = format_uptime(time.time() - START_TIME)
    text = (
        f"🤖 <b>Bot Info & Status</b>\n\n"
        f"👋 Xin chào {name_display}!\n"
        f"⏱️ Uptime: <code>{uptime_str}</code>"
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🏥 Chọn Bệnh Viện", callback_data="show_hospitals")]
    ])
    await update.message.reply_text(text, parse_mode="HTML", reply_markup=keyboard)

async def safe_edit_or_reply(
    query,
    text: str,
    reply_markup: InlineKeyboardMarkup = None,
    parse_mode: str = "HTML"
) -> None:
    """
    Chỉnh sửa nội dung tin nhắn nếu tin nhắn có text.
    Nếu tin nhắn là media/document (không có text) hoặc không thể edit,
    sẽ gửi tin nhắn mới (reply_text).
    """
    if query.message and query.message.text is not None:
        try:
            await query.edit_message_text(
                text=text,
                parse_mode=parse_mode,
                reply_markup=reply_markup
            )
            return
        except Exception as e:
            log.debug(f"Không thể edit_message_text ({e}), fallback sang reply_text")

    await query.message.reply_text(
        text=text,
        parse_mode=parse_mode,
        reply_markup=reply_markup
    )

async def show_hospitals_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Hiển thị danh sách bệnh viện dưới dạng inline button (dùng callback_data + index)."""
    query = update.callback_query
    await query.answer()

    json_path = pathlib.Path(__file__).parent / "hospitals.json"
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        hospitals = data.get("hospitals", [])
    except Exception as e:
        await query.message.reply_text(f"❌ Lỗi đọc dữ liệu: {e}")
        return

    # Xếp 2 nút mỗi hàng, dùng index làm callback_data
    rows = []
    for i in range(0, len(hospitals), 2):
        row = []
        for idx in range(i, min(i + 2, len(hospitals))):
            h = hospitals[idx]
            row.append(InlineKeyboardButton(h["name"], callback_data=f"bv:{idx}"))
        rows.append(row)

    keyboard = InlineKeyboardMarkup(rows)
    await safe_edit_or_reply(
        query=query,
        text=(
            f"🏥 <b>Danh sách {len(hospitals)} Bệnh Viện VnCare</b>\n"
            f"Chọn bệnh viện để đăng nhập:"
        ),
        reply_markup=keyboard,
    )

async def hospital_selected_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """User chọn 1 bệnh viện → yêu cầu nhập credential."""
    query = update.callback_query
    await query.answer()

    idx = int(query.data.split(":")[1])
    json_path = pathlib.Path(__file__).parent / "hospitals.json"
    with open(json_path, "r", encoding="utf-8") as f:
        hospitals = json.load(f)["hospitals"]

    hospital = hospitals[idx]
    # Lưu vào user_data để dùng ở bước tiếp theo
    context.user_data["selected_hospital"] = hospital
    context.user_data["awaiting_login"] = True

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔙 Chọn lại bệnh viện", callback_data="show_hospitals")]
    ])
    await safe_edit_or_reply(
        query=query,
        text=(
            f"🏥 Bạn đã chọn: <b>{hospital['name']}</b>\n"
            f"🔗 <code>{hospital['link']}</code>\n\n"
            f"🔑 Gửi thông tin đăng nhập theo định dạng:\n"
            f"<code>username:password</code>"
        ),
        reply_markup=keyboard,
    )

DATE_LABELS = {
    "today": "Hôm nay",
    "6": "Từ 01/01/2026 đến nay",
    "5": "Từ 01/01/2025 đến nay",
    "4": "Từ 01/01/2024 đến nay",
    "3": "Từ 01/01/2023 đến nay",
    "2": "Từ 01/01/2022 đến nay",
    "1": "Từ 01/01/2021 đến nay",
}

DATE_OPTION_BUTTONS = [
    [InlineKeyboardButton("📅 Hôm nay", callback_data="date_opt:today")],
    [
        InlineKeyboardButton("🗓️ Từ 2026 → nay", callback_data="date_opt:6"),
        InlineKeyboardButton("🗓️ Từ 2025 → nay", callback_data="date_opt:5"),
    ],
    [
        InlineKeyboardButton("🗓️ Từ 2024 → nay", callback_data="date_opt:4"),
        InlineKeyboardButton("🗓️ Từ 2023 → nay", callback_data="date_opt:3"),
    ],
    [
        InlineKeyboardButton("🗓️ Từ 2022 → nay", callback_data="date_opt:2"),
        InlineKeyboardButton("🗓️ Từ 2021 → nay", callback_data="date_opt:1"),
    ],
    [InlineKeyboardButton("🏥 Chọn lại bệnh viện", callback_data="show_hospitals")],
]

def load_hospitals_map() -> dict:
    """Trả về từ điển tra cứu bệnh viện theo domain, subdomain và tên."""
    json_path = pathlib.Path(__file__).parent / "hospitals.json"
    h_map = {}
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            for h in data.get("hospitals", []):
                name_key = h.get("name", "").strip().lower()
                domain_key = h.get("domain", "").strip().lower()
                if name_key:
                    h_map[name_key] = h
                if domain_key:
                    h_map[domain_key] = h
                    sub = domain_key.split(".")[0]
                    h_map[sub] = h
    except Exception as e:
        log.warning(f"Lỗi load hospitals.json: {e}")
    return h_map


def parse_account_line(line: str, default_hospital: dict = None, hospitals_map: dict = None) -> dict | None:
    """
    Phân tích 1 dòng chứa thông tin tài khoản.
    Hỗ trợ:
      1. https://bv... username password (hoặc : hoặc |)
      2. domain.vncare.vn username password (hoặc : hoặc |)
      3. ten_bv username password (hoặc : hoặc |)
      4. username:password hoặc username password (khi đã chọn bệnh viện trước)
    """
    line = line.strip()
    if not line or line.startswith("#") or line.startswith("//"):
        return None

    hospitals_map = hospitals_map or {}
    url, username, password = None, None, None

    # Kiểm tra phân cách bằng pipe: url|user|pass hoặc user|pass
    if "|" in line:
        pipe_parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(pipe_parts) == 3:
            first, username, password = pipe_parts[0], pipe_parts[1], pipe_parts[2]
        elif len(pipe_parts) == 2:
            if default_hospital:
                url = default_hospital.get("link")
                username, password = pipe_parts[0], pipe_parts[1]
            else:
                first, username = pipe_parts[0], pipe_parts[1]

        if not url and "first" in locals():
            if first.startswith("http://") or first.startswith("https://"):
                url = first
            elif ".vncare.vn" in first.lower():
                dom = first.lower().replace("http://", "").replace("https://", "").split("/")[0]
                url = f"https://{dom}/vnpthis"
            elif first.lower() in hospitals_map:
                url = hospitals_map[first.lower()].get("link")

    if not (url and username and password):
        parts = line.split()
        if not parts:
            return None
        first = parts[0]

        if first.startswith("http://") or first.startswith("https://"):
            url = first
            rest = line[len(first):].strip()
            if ":" in rest and len(rest.split()) == 1:
                username, _, password = rest.partition(":")
            elif len(parts) >= 3:
                username, password = parts[1], parts[2]
            elif len(parts) == 2 and ":" in parts[1]:
                username, _, password = parts[1].partition(":")
        elif ".vncare.vn" in first.lower():
            dom = first.lower().replace("http://", "").replace("https://", "").split("/")[0]
            url = f"https://{dom}/vnpthis"
            rest = line[len(first):].strip()
            if ":" in rest and len(rest.split()) == 1:
                username, _, password = rest.partition(":")
            elif len(parts) >= 3:
                username, password = parts[1], parts[2]
        elif first.lower() in hospitals_map:
            url = hospitals_map[first.lower()].get("link")
            rest = line[len(first):].strip()
            if ":" in rest and len(rest.split()) == 1:
                username, _, password = rest.partition(":")
            elif len(parts) >= 3:
                username, password = parts[1], parts[2]
        elif default_hospital:
            url = default_hospital.get("link")
            if ":" in line:
                username, _, password = line.partition(":")
            elif len(parts) >= 2:
                username, password = parts[0], parts[1]

    if not (url and username and password):
        return None

    username = username.strip()
    password = password.strip()
    if not username or not password:
        return None

    # Xác định hospital_name và display_name
    domain = url.split("://")[-1].split("/")[0]
    sub = domain.split(".")[0]
    if sub.lower() in hospitals_map:
        h_info = hospitals_map[sub.lower()]
        hospital_name = h_info.get("name", sub).lower()
        hospital_display = h_info.get("name", sub.upper())
    else:
        hospital_name = sub.lower()
        hospital_display = sub.upper()

    return {
        "url": url,
        "username": username,
        "password": password,
        "hospital_name": hospital_name,
        "display_name": hospital_display,
        "original_line": line,
    }


def parse_all_accounts(text: str, default_hospital: dict = None) -> list[dict]:
    """Phân tích văn bản thành danh sách tài khoản hợp lệ, khử trùng lặp."""
    h_map = load_hospitals_map()
    accounts = []
    seen = set()

    for line in text.strip().splitlines():
        acc = parse_account_line(line, default_hospital=default_hospital, hospitals_map=h_map)
        if acc:
            key = (acc["url"].lower(), acc["username"].lower())
            if key not in seen:
                seen.add(key)
                accounts.append(acc)

    return accounts


async def process_accounts_input(text: str, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Xử lý danh sách tài khoản (đơn hoặc đa tài khoản) gửi tới bot."""
    default_hospital = context.user_data.get("selected_hospital")
    accounts = parse_all_accounts(text, default_hospital=default_hospital)

    if not accounts:
        if context.user_data.get("awaiting_login"):
            await update.effective_message.reply_text(
                "❌ Định dạng sai. Vui lòng gửi theo các định dạng sau:\n\n"
                "• Đơn lẻ: <code>username:password</code>\n"
                "• Có link: <code>https://bv.vncare.vn/vnpthis username password</code>\n"
                "• Hàng loạt: Mỗi dòng một tài khoản hoặc gửi file <code>.txt</code>\n"
                "<i>Ví dụ:</i>\n"
                "<code>https://bvdongthap.vncare.vn/vnpthis user1 pass1\n"
                "https://bvgialai.vncare.vn/vnpthis user2 pass2</code>",
                parse_mode="HTML"
            )
        return

    context.user_data["awaiting_login"] = False

    # Xử lý 1 tài khoản
    if len(accounts) == 1:
        acc = accounts[0]
        status_msg = await update.effective_message.reply_text(
            f"⏳ Đang xác thực tài khoản <b>{acc['display_name']}</b> (<code>{acc['username']}</code>)...",
            parse_mode="HTML"
        )

        result = await login(url=acc["url"], username=acc["username"], password=acc["password"])

        if result.get("success"):
            real_base = result.get("base_url") or acc["url"]
            acc_data = {
                "url": real_base,
                "username": acc["username"],
                "password": acc["password"],
                "hospital_name": acc["hospital_name"],
                "display_name": acc["display_name"],
            }
            context.user_data["auth_sessions"] = [acc_data]
            context.user_data["auth_session"] = acc_data

            info_line = ""
            try:
                dept_res = await sel_dept(url=real_base, login_result=result, db=db)
                d_count = dept_res.get("dept_count", 0)
                r_count = dept_res.get("room_count", 0)
                if d_count or r_count:
                    info_line = f"📊 Tổng số: <b>{d_count}</b> khoa | <b>{r_count}</b> phòng\n\n"
            except Exception:
                pass

            keyboard = InlineKeyboardMarkup(DATE_OPTION_BUTTONS)
            await status_msg.edit_text(
                f"✅ <b>Đăng nhập thành công!</b>\n"
                f"🏥 Bệnh viện: <b>{acc['display_name']}</b>\n"
                f"👤 Tài khoản: <code>{acc['username']}</code>\n"
                f"{info_line}"
                f"📅 <b>Chọn khoảng thời gian xuất danh sách bệnh nhân:</b>",
                parse_mode="HTML",
                reply_markup=keyboard,
            )
            log.info(f"Login OK: {acc['display_name']} | {acc['username']} | user_id={update.effective_user.id}")
        else:
            await status_msg.edit_text(
                f"❌ <b>Đăng nhập thất bại!</b>\n"
                f"🏥 Bệnh viện: <b>{acc['display_name']}</b>\n"
                f"👤 Tài khoản: <code>{acc['username']}</code>\n"
                f"🔴 {result.get('message', 'Không đăng nhập được')}",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🏥 Chọn lại bệnh viện", callback_data="show_hospitals")]
                ])
            )
            log.warning(f"Login FAIL: {acc['display_name']} | {acc['username']} | {result.get('message')}")
        return

    # Xử lý đa tài khoản (nhiều hơn 1 tài khoản)
    total_accs = len(accounts)
    status_msg = await update.effective_message.reply_text(
        f"⏳ Đang kiểm tra và xác thực đồng thời <b>{total_accs}</b> tài khoản...",
        parse_mode="HTML"
    )

    verify_sem = asyncio.Semaphore(5)

    async def _verify(acc_item: dict) -> dict:
        async with verify_sem:
            try:
                res = await login(
                    url=acc_item["url"],
                    username=acc_item["username"],
                    password=acc_item["password"]
                )
                acc_item["success"] = res.get("success", False)
                if acc_item["success"]:
                    real_base = res.get("base_url") or acc_item["url"]
                    acc_item["url"] = real_base
                    try:
                        dept_res = await sel_dept(url=real_base, login_result=res, db=db)
                        acc_item["dept_count"] = dept_res.get("dept_count", 0)
                        acc_item["room_count"] = dept_res.get("room_count", 0)
                    except Exception:
                        acc_item["dept_count"] = 0
                        acc_item["room_count"] = 0
                else:
                    acc_item["message"] = res.get("message", "Đăng nhập thất bại")
            except Exception as e:
                acc_item["success"] = False
                acc_item["message"] = str(e)
            return acc_item

    verified_accs = await asyncio.gather(*[_verify(a) for a in accounts])
    valid_accs = [a for a in verified_accs if a.get("success")]

    lines_report = []
    for a in verified_accs:
        if a.get("success"):
            info = f"({a.get('dept_count', 0)} khoa | {a.get('room_count', 0)} phòng)"
            lines_report.append(f"✅ <b>{a['display_name']}</b> | <code>{a['username']}</code> {info}")
        else:
            msg = a.get("message", "Lỗi")
            lines_report.append(f"❌ <b>{a['display_name']}</b> | <code>{a['username']}</code>: {msg}")

    report_content = "\n".join(lines_report)

    if valid_accs:
        context.user_data["auth_sessions"] = valid_accs
        context.user_data["auth_session"] = valid_accs[0]

        keyboard = InlineKeyboardMarkup(DATE_OPTION_BUTTONS)
        await status_msg.edit_text(
            f"👥 <b>Kết quả xác thực ({len(valid_accs)}/{total_accs} tài khoản thành công):</b>\n\n"
            f"{report_content}\n\n"
            f"📅 <b>Chọn khoảng thời gian để bắt đầu xuất file cho {len(valid_accs)} tài khoản:</b>",
            parse_mode="HTML",
            reply_markup=keyboard,
        )
        log.info(f"Multi-account login: {len(valid_accs)}/{total_accs} OK for user_id={update.effective_user.id}")
    else:
        await status_msg.edit_text(
            f"❌ <b>Tất cả {total_accs} tài khoản đều xác thực thất bại:</b>\n\n"
            f"{report_content}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏥 Chọn lại bệnh viện", callback_data="show_hospitals")]
            ])
        )


async def handle_credentials(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Nhận tin nhắn văn bản chứa 1 hoặc nhiều tài khoản."""
    if not is_allowed_chat(update):
        return
    await process_accounts_input(update.message.text.strip(), update, context)


async def handle_document_accounts(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Nhận file văn bản (.txt) chứa danh sách nhiều tài khoản."""
    if not is_allowed_chat(update):
        return

    doc = update.message.document
    if not doc:
        return

    file_name = doc.file_name or ""
    if not file_name.lower().endswith(".txt"):
        return

    status_msg = await update.message.reply_text("📥 Đang tải và xử lý file danh sách tài khoản...")
    try:
        tg_file = await context.bot.get_file(doc.file_id)
        file_bytes = await tg_file.download_as_bytearray()
        content = file_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        await status_msg.edit_text(f"❌ Lỗi khi đọc file: {e}")
        return

    await status_msg.delete()
    await process_accounts_input(content, update, context)


async def export_date_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Xử lý xuất dữ liệu theo khoảng thời gian cho một hoặc nhiều tài khoản."""
    query = update.callback_query
    await query.answer()

    auth_sessions = context.user_data.get("auth_sessions")
    if not auth_sessions and context.user_data.get("auth_session"):
        auth_sessions = [context.user_data["auth_session"]]

    if not auth_sessions:
        await query.message.reply_text(
            "⚠️ Phiên làm việc đã hết hạn hoặc chưa đăng nhập. Vui lòng chọn lại bệnh viện và đăng nhập lại.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏥 Chọn Bệnh Viện", callback_data="show_hospitals")]
            ]),
            parse_mode="HTML"
        )
        return

    date_opt = query.data.split(":")[1]
    label = DATE_LABELS.get(date_opt, date_opt)

    buttons = [
        [InlineKeyboardButton("🔄 Xuất khoảng thời gian khác", callback_data="show_date_options")],
        [InlineKeyboardButton("🏥 Chọn bệnh viện khác", callback_data="show_hospitals")],
    ]

    # Trường hợp 1: Chỉ có 1 tài khoản
    if len(auth_sessions) == 1:
        session = auth_sessions[0]
        hosp_name = session.get("hospital_name", "hospital")
        disp_name = session.get("display_name", hosp_name)

        status_msg = await query.message.reply_text(
            f"⏳ <b>Đang thu thập và xuất danh sách bệnh nhân...</b>\n"
            f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
            f"📅 Khoảng thời gian: <b>{label}</b>\n"
            f"<i>Hệ thống đang tải dữ liệu song song qua Stored Procedure, vui lòng chờ...</i>",
            parse_mode="HTML"
        )

        try:
            exp_res = await export_hospital_patients_excel(
                url=session["url"],
                username=session["username"],
                password=session["password"],
                hospital_name=hosp_name,
                date_option=date_opt,
                fetch_all=True,
            )
        except Exception as e:
            log.exception(f"Lỗi khi xuất file Excel cho {disp_name}: {e}")
            await status_msg.edit_text(
                f"❌ <b>Lỗi trong quá trình xuất file:</b> <code>{e}</code>",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup(buttons)
            )
            return

        if exp_res.get("success"):
            file_path = exp_res.get("file_path")
            count = exp_res.get("count", 0)
            filename = exp_res.get("filename", f"[{hosp_name}][{count}].xlsx")

            caption = (
                f"✅ <b>Xuất file thành công!</b>\n"
                f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
                f"📅 Thời gian: <b>{label}</b>\n"
                f"👥 Tổng số bệnh nhân: <b>{count:,}</b>\n"
                f"📁 Tên file: <code>{filename}</code>"
            )

            if file_path and os.path.exists(file_path):
                with open(file_path, "rb") as doc:
                    await query.message.reply_document(
                        document=doc,
                        filename=filename,
                        caption=caption,
                        parse_mode="HTML",
                        reply_markup=InlineKeyboardMarkup(buttons),
                    )
                try:
                    os.remove(file_path)
                except Exception as e:
                    log.warning(f"Không thể xóa file tạm {file_path}: {e}")
                await status_msg.delete()
            else:
                await status_msg.edit_text(
                    f"✅ <b>Đã quét xong:</b> Tìm thấy <b>{count:,}</b> bệnh nhân ({label}).",
                    parse_mode="HTML",
                    reply_markup=InlineKeyboardMarkup(buttons),
                )
        else:
            err_msg = exp_res.get("message", "Lỗi không xác định")
            await status_msg.edit_text(
                f"❌ <b>Không thể xuất file:</b>\n{err_msg}",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup(buttons),
            )
        return

    # Trường hợp 2: Đa tài khoản (>= 2 tài khoản)
    total_accs = len(auth_sessions)
    status_msg = await query.message.reply_text(
        f"⏳ <b>Đang xuất dữ liệu hàng loạt cho {total_accs} tài khoản...</b>\n"
        f"📅 Khoảng thời gian: <b>{label}</b>\n"
        f"<i>Các file Excel hoàn tất sẽ được gửi đến bạn ngay lập tức...</i>",
        parse_mode="HTML"
    )

    hosp_name_counts = Counter(s.get("hospital_name", "hospital") for s in auth_sessions)
    completed_count = 0
    fail_count = 0
    results_summary = []

    export_sem = asyncio.Semaphore(2)

    async def _export_single(idx: int, s: dict):
        nonlocal completed_count, fail_count
        h_name = s.get("hospital_name", "hospital")
        disp = s.get("display_name", h_name)
        u_name = s.get("username", "")

        file_hosp_name = f"{h_name}_{u_name}" if hosp_name_counts[h_name] > 1 else h_name

        async with export_sem:
            try:
                exp_res = await export_hospital_patients_excel(
                    url=s["url"],
                    username=s["username"],
                    password=s["password"],
                    hospital_name=file_hosp_name,
                    date_option=date_opt,
                    fetch_all=True,
                )
                if exp_res.get("success"):
                    f_path = exp_res.get("file_path")
                    p_count = exp_res.get("count", 0)
                    fname = exp_res.get("filename", f"[{file_hosp_name}][{p_count}].xlsx")
                    completed_count += 1

                    caption = (
                        f"✅ <b>[{completed_count}/{total_accs}] Xuất file thành công!</b>\n"
                        f"🏥 Bệnh viện: <b>{disp}</b>\n"
                        f"👤 Tài khoản: <code>{u_name}</code>\n"
                        f"📅 Thời gian: <b>{label}</b>\n"
                        f"👥 Tổng số bệnh nhân: <b>{p_count:,}</b>\n"
                        f"📁 File: <code>{fname}</code>"
                    )

                    if f_path and os.path.exists(f_path):
                        with open(f_path, "rb") as doc:
                            await query.message.reply_document(
                                document=doc,
                                filename=fname,
                                caption=caption,
                                parse_mode="HTML",
                            )
                        try:
                            os.remove(f_path)
                        except Exception:
                            pass
                    results_summary.append(f"✅ <b>{disp}</b> (<code>{u_name}</code>): {p_count:,} bệnh nhân")
                else:
                    fail_count += 1
                    err = exp_res.get("message", "Lỗi xuất file")
                    results_summary.append(f"❌ <b>{disp}</b> (<code>{u_name}</code>): {err}")
            except Exception as e:
                fail_count += 1
                log.exception(f"Lỗi xuất file cho {disp} ({u_name}): {e}")
                results_summary.append(f"❌ <b>{disp}</b> (<code>{u_name}</code>): {e}")

    await asyncio.gather(*[_export_single(i, s) for i, s in enumerate(auth_sessions, start=1)])

    summary_text = (
        f"📊 <b>Tổng kết xuất file hàng loạt:</b>\n"
        f"📅 Khoảng thời gian: <b>{label}</b>\n\n"
        + "\n".join(results_summary)
        + f"\n\n🎉 <b>Hoàn tất:</b> Thành công {completed_count}/{total_accs} tài khoản."
    )
    await status_msg.edit_text(
        summary_text,
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(buttons)
    )


async def show_date_options_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Hiển thị lại menu chọn ngày cho phiên làm việc hiện tại."""
    query = update.callback_query
    await query.answer()

    auth_sessions = context.user_data.get("auth_sessions")
    if not auth_sessions and context.user_data.get("auth_session"):
        auth_sessions = [context.user_data["auth_session"]]

    if not auth_sessions:
        await query.message.reply_text(
            "⚠️ Phiên làm việc đã hết hạn. Vui lòng chọn lại bệnh viện và đăng nhập lại.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏥 Chọn Bệnh Viện", callback_data="show_hospitals")]
            ]),
            parse_mode="HTML"
        )
        return

    if len(auth_sessions) == 1:
        disp_name = auth_sessions[0].get("display_name", "Bệnh viện")
        prompt = f"📅 <b>Chọn khoảng thời gian xuất danh sách bệnh nhân cho {disp_name}:</b>"
    else:
        prompt = f"📅 <b>Chọn khoảng thời gian xuất danh sách bệnh nhân cho {len(auth_sessions)} tài khoản:</b>"

    keyboard = InlineKeyboardMarkup(DATE_OPTION_BUTTONS)
    await query.message.reply_text(
        prompt,
        parse_mode="HTML",
        reply_markup=keyboard,
    )


async def ping_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not is_allowed_chat(update):
        return
    await update.message.reply_text("🏓 Pong!")


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Lỗi xảy ra khi xử lý update: %s", context.error)


def main() -> None:
    if not config.BOT_TOKEN:
        log.error("Không tìm thấy BOT_TOKEN trong config.py!")
        return

    request = HTTPXRequest(connect_timeout=30.0, read_timeout=30.0)
    application = Application.builder().token(config.BOT_TOKEN).request(request).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("ping", ping_command))
    application.add_handler(CallbackQueryHandler(show_hospitals_callback, pattern="^show_hospitals$"))
    application.add_handler(CallbackQueryHandler(hospital_selected_callback, pattern=r"^bv:\d+$"))
    application.add_handler(CallbackQueryHandler(export_date_callback, pattern=r"^date_opt:"))
    application.add_handler(CallbackQueryHandler(show_date_options_callback, pattern="^show_date_options$"))
    application.add_handler(MessageHandler(filters.Document.FileExtension("txt"), handle_document_accounts))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_credentials))
    application.add_error_handler(error_handler)

    log.info(f"Bot đang khởi động (Chế độ Webhook: {config.USE_WEBHOOK})...")
    if config.USE_WEBHOOK:
        application.run_webhook(
            listen=config.WEBHOOK_IP,
            port=config.WEBHOOK_PORT,
            webhook_url=config.WEBHOOK_URL,
            cert=config.CERTPATH,
            allowed_updates=Update.ALL_TYPES,
        )
    else:
        application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
