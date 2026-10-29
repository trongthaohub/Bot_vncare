import asyncio
from collections import Counter
import io
import json
import logging
import logging.handlers
import re
import socket
import pathlib
import time
from urllib.parse import urlparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
import os
from telegram.request import HTTPXRequest
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from database import Database
from libs.login import login, is_retryable
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

    # Không cần chọn bệnh viện: mọi thông tin nằm trong tin nhắn người dùng gửi.
    context.user_data["awaiting_login"] = True
    for key in ("selected_hospital", "auth_session", "auth_sessions"):
        context.user_data.pop(key, None)

    # Ưu tiên hiển thị @username, nếu không có thì lấy first_name
    name_display = f"@{user.username}" if user.username else user.first_name

    uptime_str = format_uptime(time.time() - START_TIME)
    text = (
        f"🤖 <b>Bot Xuất Dữ Liệu VnCare</b>\n\n"
        f"👋 Xin chào {name_display}!\n"
        f"⏱️ Uptime: <code>{uptime_str}</code>\n\n"
        f"📌 <b>Gửi thông tin đăng nhập</b> theo một trong các định dạng sau:\n"
        f"• <code>bvabc.vncare.vn:username:password</code>\n"
        f"• <code>bvabc.vncare.vn username password</code>\n"
        f"• <code>https://bvabc.vncare.vn/vnpthis:username:password</code>\n\n"
        f"⚠️ Chỉ hỗ trợ subdomain bệnh viện của <b>vncare.vn</b> (vd: <code>bvabc.vncare.vn</code>).\n"
        f"📄 Gửi file <code>.txt</code> (mỗi dòng một tài khoản) để xuất hàng loạt."
    )
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔍 Lọc Bệnh Viện", callback_data="show_hospitals_filter")],
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
    Nếu text dài hơn giới hạn Telegram (4096 ký tự) thì tự cắt thành nhiều phần.
    """
    max_len = 3800  # chừa margin cho entity/hình thức parse
    chunks = [text[i:i + max_len] for i in range(0, len(text), max_len)] if len(text) > max_len else [text]

    if query.message and query.message.text is not None:
        try:
            await query.edit_message_text(
                text=chunks[0],
                parse_mode=parse_mode,
                reply_markup=reply_markup if len(chunks) == 1 else None
            )
            for chunk in chunks[1:]:
                await query.message.reply_text(
                    text=chunk,
                    parse_mode=parse_mode,
                    reply_markup=reply_markup if chunk is chunks[-1] else None
                )
            return
        except Exception as e:
            log.debug(f"Không thể edit_message_text ({e}), fallback sang reply_text")

    for chunk in chunks:
        await query.message.reply_text(
            text=chunk,
            parse_mode=parse_mode,
            reply_markup=reply_markup if chunk is chunks[-1] else None
        )

async def show_hospitals_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """(Không còn dùng) Trước đây hiển thị danh sách bệnh viện để chọn trước khi đăng nhập."""
    query = update.callback_query
    await query.answer()
    await safe_edit_or_reply(
        query=query,
        text=(
            "ℹ️ Không cần chọn bệnh viện nữa.\n\n"
            "Gửi trực tiếp thông tin đăng nhập theo dạng:\n"
            "<code>bvabc.vncare.vn:username:password</code>\n"
            "hoặc <code>bvabc.vncare.vn username password</code>"
        ),
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
]

# ===================== CHỨC NĂNG LỌC TÀI KHOẢN THEO BỆNH VIỆN =====================
# Luồng: /start → "🔍 Lọc Bệnh Viện" → chọn bệnh viện → nhập số lượng cần lấy
#        → gửi file .txt (định dạng username:password) → bot đăng nhập thử
#        → gửi lại file .txt chỉ chứa các tài khoản đăng nhập thành công.

FILTER_CONCURRENCY = 10
# Số lần thử lại khi gặp lỗi TẠM THỜI (mạng/timeout/phản hồi lạ).
# Sai mật khẩu thì không thử lại vì thử cũng vô nghĩa.
FILTER_RETRY = 3
FILTER_RETRY_DELAY = 1.5   # giây, tăng dần theo số lần thử
# Khoảng cách tối thiểu giữa 2 lần cập nhật tiến độ (giây).
# Cập nhật mỗi batch (~1.5 edit/s) sẽ bị Telegram trả 429 liên tục và làm
# chậm cả luồng vì mỗi batch phải chờ xong edit mới chạy tiếp.
FILTER_EDIT_INTERVAL = 3.0


async def show_hospitals_filter_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Hiển thị danh sách bệnh viện để chọn bệnh viện cần lọc tài khoản."""
    query = update.callback_query
    await query.answer()

    json_path = pathlib.Path(__file__).parent / "hospitals.json"
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            hospitals = json.load(f).get("hospitals", [])
    except Exception as e:
        await query.message.reply_text(f"❌ Lỗi đọc dữ liệu: {e}")
        return

    if not hospitals:
        await query.message.reply_text("❌ Danh sách bệnh viện trống.")
        return

    rows = []
    for i in range(0, len(hospitals), 2):
        row = []
        for idx in range(i, min(i + 2, len(hospitals))):
            row.append(InlineKeyboardButton(hospitals[idx]["name"], callback_data=f"flt:{idx}"))
        rows.append(row)
    rows.append([InlineKeyboardButton("🔙 Huỷ", callback_data="filter_cancel")])

    await safe_edit_or_reply(
        query=query,
        text=(
            f"🔍 <b>Lọc tài khoản theo bệnh viện</b>\n\n"
            f"Chọn bệnh viện bạn muốn lọc ({len(hospitals)} bệnh viện):"
        ),
        reply_markup=InlineKeyboardMarkup(rows),
    )


async def filter_hospital_selected_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """User chọn bệnh viện để lọc → hỏi số lượng tài khoản cần lấy."""
    query = update.callback_query
    await query.answer()

    try:
        idx = int(query.data.split(":")[1])
        json_path = pathlib.Path(__file__).parent / "hospitals.json"
        with open(json_path, "r", encoding="utf-8") as f:
            hospitals = json.load(f)["hospitals"]
        hospital = hospitals[idx]
    except Exception as e:
        await query.message.reply_text(f"❌ Không đọc được bệnh viện: {e}")
        return

    # Lưu trạng thái lọc, tắt trạng thái đăng nhập thường
    context.user_data["filter_hospital"] = hospital
    context.user_data["awaiting_filter_count"] = True
    context.user_data["awaiting_filter_file"] = False
    context.user_data.pop("filter_target_count", None)
    context.user_data["awaiting_login"] = False

    await safe_edit_or_reply(
        query=query,
        text=(
            f"🏥 Bệnh viện lọc: <b>{hospital['name']}</b>\n"
            f"🔗 <code>{hospital['link']}</code>\n\n"
            f"🔢 Bạn muốn lấy bao nhiêu tài khoản hoạt động?\n"
            f"<i>Nhập một số, ví dụ: <code>10</code></i>"
        ),
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔙 Huỷ", callback_data="filter_cancel")]
        ]),
    )


async def handle_filter_count(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Nhận số lượng tài khoản cần lọc và yêu cầu gửi file .txt."""
    text = (update.message.text or "").strip()
    if not text.isdigit() or int(text) <= 0:
        await update.message.reply_text(
            "❌ Vui lòng nhập một số nguyên dương (ví dụ: <code>10</code>).",
            parse_mode="HTML",
        )
        return

    target = int(text)
    hospital = context.user_data.get("filter_hospital") or {}
    context.user_data["filter_target_count"] = target
    context.user_data["awaiting_filter_count"] = False
    context.user_data["awaiting_filter_file"] = True

    await update.message.reply_text(
        f"✅ Sẽ lọc lấy <b>{target}</b> tài khoản hoạt động từ <b>{hospital.get('name', '?')}</b>.\n\n"
        f"📄 Gửi file <code>.txt</code> chứa danh sách tài khoản, mỗi dòng một tài khoản:\n"
        f"<code>username:password</code>\n\n"
        f"⚠️ Số tài khoản trong file phải <b>lớn hơn {target}</b>.",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔙 Huỷ", callback_data="filter_cancel")]
        ]),
    )


async def process_filter_file(
    content: str,
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    status_msg,
) -> None:
    """Lọc file tài khoản: đăng nhập thử và chỉ giữ lại tài khoản thành công."""
    hospital = context.user_data.get("filter_hospital") or {}
    target = context.user_data.get("filter_target_count")

    if not hospital or not target:
        await status_msg.edit_text(
            "⚠️ Phiên lọc đã hết hạn. Vui lòng chọn lại bệnh viện để lọc.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔍 Lọc Bệnh Viện", callback_data="show_hospitals_filter")]
            ]),
        )
        return

    accounts = parse_all_accounts(content, default_hospital=hospital)
    total = len(accounts)

    if total <= target:
        await status_msg.edit_text(
            f"❌ File chỉ có <b>{total}</b> tài khoản hợp lệ.\n"
            f"Số lượng trong file phải <b>lớn hơn {target}</b>.\n\n"
            f"Vui lòng gửi file khác.",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔙 Huỷ", callback_data="filter_cancel")]
            ]),
        )
        return

    await status_msg.edit_text(
        f"⏳ <b>Đang kiểm tra {total} tài khoản...</b>\n"
        f"🏥 Bệnh viện: <b>{hospital.get('name', '?')}</b>\n"
        f"🎯 Cần lấy: <b>{target}</b> tài khoản hoạt động\n"
        f"<i>Đã kiểm tra 0/{total} — tìm được 0/{target}</i>",
        parse_mode="HTML",
    )

    sem = asyncio.Semaphore(FILTER_CONCURRENCY)
    valid: list[dict] = []
    checked = 0
    net_failed: list[dict] = []   # lỗi tạm thời (mạng/chập chờn)
    last_edit = 0.0               # lần cuối cập nhật tiến độ

    async def _check(acc: dict) -> str:
        """Trả về "ok" | "auth" | "network".

        Thử lại các lỗi TẠM THỜI để không loại oan tài khoản đúng khi mạng
        chập. Chỉ kết luận thất bại khi server trả lời rõ ràng là sai
        tài khoản/mật khẩu ("auth").
        """
        async with sem:
            last_kind = "network"
            for attempt in range(1, FILTER_RETRY + 1):
                try:
                    res = await login(
                        url=acc["url"], username=acc["username"], password=acc["password"]
                    )
                except Exception as e:
                    log.debug(f"Lỗi kiểm tra {acc.get('username')}: {e}")
                    res = {"success": False, "error_kind": "network", "message": str(e)}

                kind = res.get("error_kind")
                if res.get("success"):
                    return "ok"
                if kind == "auth":
                    # Sai mật khẩu -> thử lại vô nghĩa
                    return "auth"

                last_kind = kind or "network"
                if attempt < FILTER_RETRY:
                    await asyncio.sleep(FILTER_RETRY_DELAY * attempt)

            return last_kind

    # Kiểm tra theo từng đợt, dừng ngay khi đủ số lượng cần lấy
    for start in range(0, total, FILTER_CONCURRENCY):
        batch = accounts[start:start + FILTER_CONCURRENCY]
        results = await asyncio.gather(*[_check(a) for a in batch])
        for acc, kind in zip(batch, results):
            if kind == "ok":
                acc["success"] = True
                valid.append(acc)
            elif kind == "network":
                acc["error_kind"] = "network"
                net_failed.append(acc)
        checked += len(batch)

        # Chỉ cập nhật tiến độ tối đa 1 lần / FILTER_EDIT_INTERVAL giây.
        # Edit mỗi batch vừa gây 429 vừa chặn vòng lặp chờ Telegram trả lời.
        now = time.monotonic()
        if now - last_edit >= FILTER_EDIT_INTERVAL:
            last_edit = now
            try:
                await status_msg.edit_text(
                    f"⏳ <b>Đang kiểm tra tài khoản...</b>\n"
                    f"🏥 Bệnh viện: <b>{hospital.get('name', '?')}</b>\n"
                    f"🎯 Cần lấy: <b>{target}</b> tài khoản hoạt động\n"
                    f"<i>Đã kiểm tra {checked}/{total} — tìm được {len(valid)}/{target}</i>",
                    parse_mode="HTML",
                )
            except Exception:
                pass

        if len(valid) >= target:
            break

    context.user_data["awaiting_filter_file"] = False
    context.user_data.pop("filter_target_count", None)

    if not valid:
        # Phân biệt "sai hết mật khẩu" với "lỗi hạ tầng": nếu phần lớn là lỗi
        # mạng thì đừng nói với người dùng rằng tài khoản của họ sai.
        if net_failed and len(net_failed) >= max(1, checked * 0.5):
            await status_msg.edit_text(
                f"⚠️ <b>Không kiểm tra được tài khoản do lỗi kết nối.</b>\n"
                f"🏥 Bệnh viện: <b>{hospital.get('name', '?')}</b>\n"
                f"👥 Đã kiểm tra: <b>{checked}</b> — lỗi mạng: <b>{len(net_failed)}</b>\n\n"
                f"<i>Tài khoản chưa chắc sai. Vui lòng thử lọc lại sau ít phút.</i>",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔍 Lọc lại", callback_data="show_hospitals_filter")]
                ]),
            )
        else:
            extra = f"\n⚠️ Lỗi mạng (chưa kết luận): <b>{len(net_failed)}</b>" if net_failed else ""
            await status_msg.edit_text(
                f"❌ <b>Không có tài khoản nào đăng nhập thành công.</b>\n"
                f"🏥 Bệnh viện: <b>{hospital.get('name', '?')}</b>\n"
                f"👥 Đã kiểm tra: <b>{checked}</b> tài khoản{extra}",
                parse_mode="HTML",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔍 Lọc lại", callback_data="show_hospitals_filter")]
                ]),
            )
        return

    final_accs = valid[:target]
    payload = "\n".join(f"{a['username']}:{a['password']}" for a in final_accs)
    hospital_name = hospital.get("name", "hospital")
    filename = f"[{hospital_name}][loc][{len(final_accs)}].txt"

    caption = (
        f"✅ <b>Lọc tài khoản hoàn tất!</b>\n"
        f"🏥 Bệnh viện: <b>{hospital_name}</b>\n"
        f"📥 Tài khoản nhận vào: <b>{total}</b>\n"
        f"🔍 Đã kiểm tra: <b>{checked}</b>\n"
        f"🎯 Lấy được: <b>{len(final_accs)}/{target}</b> tài khoản hoạt động\n"
        f"📁 File: <code>{filename}</code>"
        + (f"\n⚠️ <b>{len(net_failed)}</b> tài khoản lỗi mạng, chưa kết luận được"
           if net_failed else "")
    )

    try:
        await status_msg.reply_document(
            document=io.BytesIO(payload.encode("utf-8")),
            filename=filename,
            caption=caption,
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔍 Lọc tiếp", callback_data="show_hospitals_filter")]
            ]),
        )
    except Exception as e:
        log.exception(f"Lỗi gửi file kết quả lọc: {e}")
        await status_msg.edit_text(f"❌ Không gửi được file kết quả: <code>{e}</code>", parse_mode="HTML")
        return

    await status_msg.delete()
    log.info(
        f"Filter OK: {hospital_name} | lấy {len(final_accs)}/{target} "
        f"(kiểm tra {checked}/{total}) | user_id={update.effective_user.id}"
    )


async def filter_cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Huỷ phiên lọc tài khoản."""
    query = update.callback_query
    await query.answer()

    for key in (
        "filter_hospital",
        "filter_target_count",
        "awaiting_filter_count",
        "awaiting_filter_file",
    ):
        context.user_data.pop(key, None)

    await safe_edit_or_reply(
        query=query,
        text="🚫 Đã huỷ phiên lọc tài khoản.\n\nGửi thông tin đăng nhập dạng <code>bvabc.vncare.vn:username:password</code> để xuất dữ liệu.",
    )


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


VNCARE_HOST_RE = re.compile(r"^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+vncare\.vn$")


def normalize_vncare_url(value: str) -> str | None:
    """Chuẩn hóa URL về dạng https://<subdomain>.vncare.vn/vnpthis.

    Chỉ chấp nhận host là subdomain bệnh viện của vncare.vn (vd: bvabc.vncare.vn).
    Trả về None nếu không hợp lệ.
    """
    raw = (value or "").strip().strip("'\"<>")
    if not raw:
        return None
    if "://" not in raw:
        raw = "https://" + raw
    try:
        host = (urlparse(raw).hostname or "").strip().lower().rstrip(".")
    except ValueError:
        return None
    if not VNCARE_HOST_RE.match(host):
        return None
    # Mọi bệnh viện VnCare dùng chung backend /vnpthis.
    return f"https://{host}/vnpthis"


def _split_url_user_pass(line: str) -> tuple[str, str, str] | None:
    """Tách dòng thành (url, username, password) mà không cần biết URL đúng/sai."""
    if "|" in line:
        parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(parts) >= 3:
            return parts[0], parts[1], "|".join(parts[2:])
        return None

    tokens = line.split()
    if len(tokens) >= 3:
        return tokens[0], tokens[1], " ".join(tokens[2:])
    if len(tokens) == 2:
        user, sep, pwd = tokens[1].partition(":")
        if sep and user and pwd:
            return tokens[0], user, pwd
        return None
    if len(tokens) == 1:
        raw = tokens[0]
        if "://" in raw:
            scheme, rest = raw.split("://", 1)
            parts = rest.split(":")
            if len(parts) >= 3:
                return f"{scheme}://{parts[0]}", parts[1], ":".join(parts[2:])
        else:
            parts = raw.split(":")
            if len(parts) >= 3:
                return parts[0], parts[1], ":".join(parts[2:])
    return None


def parse_account_line(line: str, default_hospital: dict = None, hospitals_map: dict = None) -> dict | None:
    """
    Phân tích 1 dòng thông tin đăng nhập.

    Định dạng chính (URL bắt buộc là subdomain vncare.vn):
      1. bvabc.vncare.vn:username:password
      2. bvabc.vncare.vn username password
      3. https://bvabc.vncare.vn/vnpthis:username:password
      4. url|username|password
    Ngoại lệ: nếu default_hospital được truyền (luồng lọc tài khoản) thì
    username:password / username password vẫn hợp lệ vì đã biết bệnh viện.
    """
    # File CSV / copy tu bang tinh thuong dinh dau phay o dau-cuoi dong.
    # Neu khong bo, dau phay se lot vao username -> login luon that bai
    # (trieu chung: tai khoan dau OK, cac tai khoan sau toan fail).
    line = line.strip().strip(",;").strip()
    if not line or line.startswith("#") or line.startswith("//"):
        return None

    hospitals_map = hospitals_map or {}
    url, username, password = None, None, None

    split = _split_url_user_pass(line)
    if split:
        raw_url, split_user, split_pass = split
        url = normalize_vncare_url(raw_url)
        if not url:
            # Cho phép gửi tên bệnh viện ngắn (vd: bvlaichau) nếu có trong danh mục
            entry = hospitals_map.get(raw_url.strip().lower())
            if entry:
                url = normalize_vncare_url(entry.get("link", ""))
        if url:
            username, password = split_user, split_pass
        elif default_hospital:
            # Luồng lọc: bệnh viện đã biết, dòng chỉ là "user pass ..." nên
            # phần đầu dòng thực ra là username chứ không phải URL.
            url = normalize_vncare_url(default_hospital.get("link", "")) or default_hospital.get("link")
            username, password = raw_url, (" ".join(line.split()[1:]) or split_pass)
        else:
            return None
    elif default_hospital:
        url = normalize_vncare_url(default_hospital.get("link", "")) or default_hospital.get("link")
        if "|" in line:
            pipe_parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(pipe_parts) < 2:
                return None
            username, password = pipe_parts[0], "|".join(pipe_parts[1:])
        else:
            tokens = line.split()
            if len(tokens) >= 2:
                username, password = tokens[0], " ".join(tokens[1:])
            else:
                username, _, password = line.partition(":")
    else:
        return None

    username = (username or "").strip()
    password = (password or "").strip()
    if not (url and username and password):
        return None

    # Xác định hospital_name và display_name từ host
    host = url.split("://")[-1].split("/")[0].lower()
    sub = host.split(".")[0]
    entry = hospitals_map.get(host) or hospitals_map.get(sub)
    if entry:
        hospital_name = entry.get("name", sub).lower()
        hospital_display = entry.get("name", sub.upper())
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


# Cho phép dán nhiều tài khoản trên 1 dòng, ngăn cách bằng dấu phẩy/chấm phẩy
# (chỉ tách khi phần tiếp theo bắt đầu bằng URL, tránh cắt nhầm mật khẩu).
_ENTRY_SPLIT_RE = re.compile(
    r"\s*[,;]\s*(?=(?:https?://|[a-z0-9][a-z0-9.-]*\.vncare\.vn))",
    re.IGNORECASE,
)


def parse_all_accounts(text: str, default_hospital: dict = None) -> list[dict]:
    """Phân tích văn bản thành danh sách tài khoản hợp lệ, khử trùng lặp."""
    h_map = load_hospitals_map()
    accounts = []
    seen = set()

    for raw_line in text.strip().splitlines():
        for line in _ENTRY_SPLIT_RE.split(raw_line):
            acc = parse_account_line(line, default_hospital=default_hospital, hospitals_map=h_map)
            if acc:
                key = (acc["url"].lower(), acc["username"].lower())
                if key not in seen:
                    seen.add(key)
                    accounts.append(acc)
            elif line.strip() and not line.strip().startswith(("#", "//")):
                log.debug(f"Bỏ qua dòng không hợp lệ (thiếu URL vncare hoặc sai định dạng): {line.strip()[:80]}")

    return accounts


async def process_accounts_input(text: str, update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Xử lý danh sách tài khoản (đơn hoặc đa tài khoản) gửi tới bot."""
    if not text:
        return
    # Không cần chọn bệnh viện: URL phải nằm ngay trong từng dòng.
    accounts = parse_all_accounts(text)

    if not accounts:
        if context.user_data.get("awaiting_login"):
            await update.effective_message.reply_text(
                "❌ Không nhận diện được thông tin đăng nhập.\n\n"
                "Gửi theo một trong các định dạng:\n"
                "• <code>bvabc.vncare.vn:username:password</code>\n"
                "• <code>bvabc.vncare.vn username password</code>\n"
                "• <code>https://bvabc.vncare.vn/vnpthis:username:password</code>\n"
                "• Hàng loạt: mỗi dòng một tài khoản hoặc gửi file <code>.txt</code>\n\n"
                "⚠️ URL phải là subdomain bệnh viện của <b>vncare.vn</b> (vd: <code>bvabc.vncare.vn</code>).",
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
            # login() có thể trả về base_url sau redirect; chỉ dùng nếu vẫn là host vncare.
            real_base = normalize_vncare_url(result.get("base_url") or "") or acc["url"]
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
            # Phân biệt sai mật khẩu với lỗi mạng để không nói sai với người dùng.
            if is_retryable(result):
                title = "⚠️ <b>Không kết nối được tới bệnh viện!</b>"
                hint = "\n<i>Tài khoản chưa chắc sai. Vui lòng thử lại sau ít phút.</i>"
            else:
                title = "❌ <b>Đăng nhập thất bại!</b>"
                hint = ""
            await status_msg.edit_text(
                f"{title}\n"
                f"🏥 Bệnh viện: <b>{acc['display_name']}</b>\n"
                f"👤 Tài khoản: <code>{acc['username']}</code>\n"
                f"🔴 {result.get('message', 'Không đăng nhập được')}"
                f"{hint}",
                parse_mode="HTML",
            )
            log.warning(
                f"Login FAIL: {acc['display_name']} | {acc['username']} | "
                f"kind={result.get('error_kind')} | {result.get('message')}"
            )
        return

    # Xử lý đa tài khoản (nhiều hơn 1 tài khoản)
    total_accs = len(accounts)
    status_msg = await update.effective_message.reply_text(
        f"⏳ Đang kiểm tra và xác thực đồng thời <b>{total_accs}</b> tài khoản...",
        parse_mode="HTML"
    )

    verify_sem = asyncio.Semaphore(5)

    async def _verify(acc_item: dict) -> dict:
        """Xác thực 1 tài khoản. Thử lại khi gặp lỗi TẠM THỜI (mạng)."""
        async with verify_sem:
            res = {"success": False, "error_kind": "network", "message": "Chưa chạy"}
            for attempt in range(1, FILTER_RETRY + 1):
                try:
                    res = await login(
                        url=acc_item["url"],
                        username=acc_item["username"],
                        password=acc_item["password"]
                    )
                except Exception as e:
                    res = {"success": False, "error_kind": "network", "message": str(e)}

                if res.get("success") or res.get("error_kind") == "auth":
                    break
                if attempt < FILTER_RETRY:
                    await asyncio.sleep(FILTER_RETRY_DELAY * attempt)

            acc_item["success"] = res.get("success", False)
            acc_item["error_kind"] = res.get("error_kind")
            if acc_item["success"]:
                real_base = normalize_vncare_url(res.get("base_url") or "") or acc_item["url"]
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
            # Lỗi mạng dùng dấu ⚠️ để không bị hiểu là sai mật khẩu.
            mark = "⚠️" if a.get("error_kind") == "network" else "❌"
            lines_report.append(
                f"{mark} <b>{a['display_name']}</b> | <code>{a['username']}</code>: {msg}"
            )

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
            f"{report_content}\n\n"
            f"<i>Kiểm tra lại URL (phải là subdomain vncare.vn) và thông tin đăng nhập.</i>",
            parse_mode="HTML",
        )


async def handle_credentials(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Nhận tin nhắn văn bản chứa 1 hoặc nhiều tài khoản (kèm URL vncare)."""
    if not is_allowed_chat(update):
        return

    # Đang trong luồng lọc tài khoản → chờ số lượng cần lấy
    if context.user_data.get("awaiting_filter_count"):
        await handle_filter_count(update, context)
        return

    context.user_data["awaiting_login"] = True
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

    is_filter_flow = bool(context.user_data.get("awaiting_filter_file"))

    status_msg = await update.message.reply_text(
        "📥 Đang tải file danh sách tài khoản..." if is_filter_flow
        else "📥 Đang tải và xử lý file danh sách tài khoản..."
    )
    try:
        tg_file = await context.bot.get_file(doc.file_id)
        file_bytes = await tg_file.download_as_bytearray()
        content = file_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        await status_msg.edit_text(f"❌ Lỗi khi đọc file: {e}")
        return

    if is_filter_flow:
        await process_filter_file(content, update, context, status_msg)
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
            "⚠️ Phiên làm việc đã hết hạn hoặc chưa đăng nhập.\n"
            "Vui lòng gửi lại thông tin đăng nhập dạng <code>bvabc.vncare.vn:username:password</code>.",
            parse_mode="HTML"
        )
        return

    # Chỉ xuất dữ liệu khi URL vẫn là subdomain bệnh viện của vncare.vn.
    vncare_sessions = [s for s in auth_sessions if normalize_vncare_url(s.get("url") or "")]
    if not vncare_sessions:
        await query.message.reply_text(
            "❌ URL phiên đăng nhập không phải subdomain của <b>vncare.vn</b> nên không thể xuất dữ liệu.\n"
            "Vui lòng gửi lại thông tin đăng nhập.",
            parse_mode="HTML"
        )
        return
    auth_sessions = vncare_sessions

    date_opt = query.data.split(":")[1]
    label = DATE_LABELS.get(date_opt, date_opt)

    buttons = [
        [InlineKeyboardButton("🔄 Xuất khoảng thời gian khác", callback_data="show_date_options")],
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

        last_edit = [0.0]
        async def on_progress(progress_text: str):
            now = time.time()
            if now - last_edit[0] >= 1.5:
                last_edit[0] = now
                try:
                    await status_msg.edit_text(
                        f"⏳ <b>Đang thu thập và xuất danh sách bệnh nhân...</b>\n"
                        f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
                        f"📅 Khoảng thời gian: <b>{label}</b>\n\n"
                        f"{progress_text}",
                        parse_mode="HTML"
                    )
                except Exception:
                    pass

        try:
            exp_res = await export_hospital_patients_excel(
                url=session["url"],
                username=session["username"],
                password=session["password"],
                hospital_name=hosp_name,
                date_option=date_opt,
                fetch_all=True,
                progress_callback=on_progress,
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
            raw_file_path = exp_res.get("raw_file_path") or file_path
            count = exp_res.get("count", 0)
            filename = exp_res.get("filename", f"[{hosp_name}][{count}].xlsx")
            is_zip = exp_res.get("is_zip", False)

            f_size_bytes = os.path.getsize(file_path) if file_path and os.path.exists(file_path) else 0
            f_size_mb = f_size_bytes / (1024 * 1024)

            size_desc = f"💾 Dung lượng: <b>{f_size_mb:.1f} MB</b>"
            if is_zip:
                orig_mb = exp_res.get("original_size", 0) / (1024 * 1024)
                size_desc += f" <i>(đã nén từ {orig_mb:.1f} MB để tải nhanh)</i>"

            caption = (
                f"✅ <b>Xuất file thành công!</b>\n"
                f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
                f"📅 Thời gian: <b>{label}</b>\n"
                f"👥 Tổng số bệnh nhân: <b>{count:,}</b>\n"
                f"📁 Tên file: <code>{filename}</code>\n"
                f"{size_desc}"
            )

            if file_path and os.path.exists(file_path):
                # Khởi động background heartbeat cập nhật tiến trình truyền tải theo thời gian thực
                stop_heartbeat = asyncio.Event()

                async def _upload_heartbeat():
                    start_t = time.time()
                    while not stop_heartbeat.is_set():
                        try:
                            await asyncio.wait_for(stop_heartbeat.wait(), timeout=7.0)
                            break
                        except asyncio.TimeoutError:
                            elapsed = int(time.time() - start_t)
                            mins, secs = divmod(elapsed, 60)
                            el_str = f"{mins}m {secs}s" if mins > 0 else f"{secs}s"
                            try:
                                await status_msg.edit_text(
                                    f"📤 <b>Đang truyền file lên Telegram...</b>\n"
                                    f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
                                    f"📁 File: <code>{filename}</code> ({count:,} bệnh nhân)\n"
                                    f"{size_desc}\n"
                                    f"⏱️ Đang truyền dữ liệu: <b>{el_str}</b>\n"
                                    f"<i>Hệ thống đang gửi dữ liệu qua mạng, vui lòng chờ...</i>",
                                    parse_mode="HTML"
                                )
                            except Exception:
                                pass

                heartbeat_task = asyncio.create_task(_upload_heartbeat())

                sent_ok = False
                for attempt in range(1, 4):
                    try:
                        with open(file_path, "rb") as doc:
                            await query.message.reply_document(
                                document=doc,
                                filename=filename,
                                caption=caption,
                                parse_mode="HTML",
                                reply_markup=InlineKeyboardMarkup(buttons),
                                read_timeout=600.0,
                                write_timeout=600.0,
                                connect_timeout=60.0,
                            )
                        sent_ok = True
                        break
                    except Exception as upload_err:
                        log.warning(f"Lần gửi file thứ {attempt}/3 thất bại ({upload_err})")
                        if attempt < 3:
                            await asyncio.sleep(2)

                stop_heartbeat.set()
                try:
                    await heartbeat_task
                except Exception:
                    pass

                if sent_ok:
                    for p in set([file_path, raw_file_path]):
                        try:
                            if p and os.path.exists(p):
                                os.remove(p)
                        except Exception as e:
                            log.warning(f"Không thể xóa file tạm {p}: {e}")
                    try:
                        await status_msg.delete()
                    except Exception:
                        pass
                else:
                    await status_msg.edit_text(
                        f"⚠️ <b>Lỗi mạng Telegram khi tải file lên (Timed out).</b>\n"
                        f"📁 File đã được tạo thành công trên VPS:\n"
                        f"<code>{file_path}</code> ({count:,} bệnh nhân)\n\n"
                        f"<i>Bạn có thể lấy trực tiếp file trên VPS hoặc bấm thử lại bên dưới.</i>",
                        parse_mode="HTML",
                        reply_markup=InlineKeyboardMarkup(buttons),
                    )
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
                    raw_f_path = exp_res.get("raw_file_path") or f_path
                    is_z = exp_res.get("is_zip", False)
                    completed_count += 1

                    f_size_b = os.path.getsize(f_path) if f_path and os.path.exists(f_path) else 0
                    f_size_m = f_size_b / (1024 * 1024)
                    size_s = f"💾 Dung lượng: <b>{f_size_m:.1f} MB</b>"
                    if is_z:
                        o_m = exp_res.get("original_size", 0) / (1024 * 1024)
                        size_s += f" <i>(đã nén từ {o_m:.1f} MB)</i>"

                    caption = (
                        f"✅ <b>[{completed_count}/{total_accs}] Xuất file thành công!</b>\n"
                        f"🏥 Bệnh viện: <b>{disp}</b>\n"
                        f"👤 Tài khoản: <code>{u_name}</code>\n"
                        f"📅 Thời gian: <b>{label}</b>\n"
                        f"👥 Tổng số bệnh nhân: <b>{p_count:,}</b>\n"
                        f"📁 File: <code>{fname}</code>\n"
                        f"{size_s}"
                    )

                    if f_path and os.path.exists(f_path):
                        for attempt in range(1, 4):
                            try:
                                with open(f_path, "rb") as doc:
                                    await query.message.reply_document(
                                        document=doc,
                                        filename=fname,
                                        caption=caption,
                                        parse_mode="HTML",
                                        read_timeout=600.0,
                                        write_timeout=600.0,
                                        connect_timeout=60.0,
                                    )
                                break
                            except Exception as ex:
                                log.warning(f"Lần gửi file {fname} thứ {attempt} lỗi: {ex}")
                                if attempt < 3:
                                    await asyncio.sleep(2)
                        for p in set([f_path, raw_f_path]):
                            try:
                                if p and os.path.exists(p):
                                    os.remove(p)
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
            "⚠️ Phiên làm việc đã hết hạn.\n"
            "Vui lòng gửi lại thông tin đăng nhập dạng <code>bvabc.vncare.vn:username:password</code>.",
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

    # Chống chạy đôi: giữ một socket cục bộ cho đến khi tiến trình thoát.
    _single_instance = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        _single_instance.bind(("127.0.0.1", 47615))
    except OSError:
        raise SystemExit("Bot instance khác đang chạy rồi (port 47615 bị chiếm). Thoát để trùng getUpdates -> 409.")

    # trust_env=False: bỏ qua biến proxy của môi trường (Windows registry / *_proxy).
    # Có proxy MITM cục bộ (Reqable/9router) thì TLS tới api.telegram.org sẽ fail
    # với CERTIFICATE_VERIFY_FAILED, hoặc chết hẳn khi proxy không chạy.
    request = HTTPXRequest(
        connect_timeout=60.0,
        read_timeout=60.0,
        write_timeout=600.0,
        media_write_timeout=600.0,
        httpx_kwargs={"trust_env": False},
    )
    application = Application.builder().token(config.BOT_TOKEN).request(request).build()

    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("ping", ping_command))
    application.add_handler(CallbackQueryHandler(show_hospitals_callback, pattern="^show_hospitals$"))
    application.add_handler(CallbackQueryHandler(show_hospitals_filter_callback, pattern="^show_hospitals_filter$"))
    application.add_handler(CallbackQueryHandler(filter_hospital_selected_callback, pattern=r"^flt:\d+$"))
    application.add_handler(CallbackQueryHandler(filter_cancel_callback, pattern="^filter_cancel$"))
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
