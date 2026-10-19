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

    # Xẽp 2 nút mỗi hàng, dùng index làm callback_data
    rows = []
    for i in range(0, len(hospitals), 2):
        row = []
        for idx in range(i, min(i + 2, len(hospitals))):
            h = hospitals[idx]
            row.append(InlineKeyboardButton(h["name"], callback_data=f"bv:{idx}"))
        rows.append(row)

    keyboard = InlineKeyboardMarkup(rows)
    await query.edit_message_text(
        f"🏥 <b>Danh sách {len(hospitals)} Bệnh Viện VnCare</b>\n"
        f"Chọn bệnh viện để đăng nhập:",
        parse_mode="HTML",
        reply_markup=keyboard
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
    await query.edit_message_text(
        f"🏥 Bạn đã chọn: <b>{hospital['name']}</b>\n"
        f"🔗 <code>{hospital['link']}</code>\n\n"
        f"🔑 Gửi thông tin đăng nhập theo định dạng:\n"
        f"<code>username:password</code>",
        parse_mode="HTML",
        reply_markup=keyboard
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

async def handle_credentials(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Nhận thông tin đăng nhập từ người dùng.
    Hỗ trợ 2 hình thức:
    1. Đã chọn bệnh viện trước: gửi 'username:password' hoặc 'username password'.
    2. Gửi trực tiếp link bệnh viện: 'https://bvdongthap.vncare.vn/vnpthis username password'.
    """
    text = update.message.text.strip()

    url = None
    username = None
    password = None
    hospital_name = None
    hospital_display = None

    tokens = text.split()
    if tokens and (tokens[0].startswith("http://") or tokens[0].startswith("https://")):
        url = tokens[0]
        if len(tokens) >= 3:
            username = tokens[1]
            password = tokens[2]
        elif len(tokens) == 2 and ":" in tokens[1]:
            username, _, password = tokens[1].partition(":")

        domain = url.split("://")[-1].split("/")[0]
        hospital_name = domain.split(".")[0]
        hospital_display = hospital_name.upper()
    elif context.user_data.get("awaiting_login"):
        hospital = context.user_data.get("selected_hospital", {})
        url = hospital.get("link")
        hospital_name = hospital.get("name", "hospital").lower()
        hospital_display = hospital.get("name", "Bệnh viện")

        if ":" in text:
            username, _, password = text.partition(":")
        elif len(tokens) >= 2:
            username = tokens[0]
            password = tokens[1]
    else:
        return

    if not url or not username or not password:
        await update.message.reply_text(
            "❌ Định dạng sai. Vui lòng gửi theo dạng:\n"
            "<code>username:password</code> (nếu đã chọn bệnh viện)\n"
            "hoặc: <code>https://bv.vncare.vn/vnpthis username password</code>",
            parse_mode="HTML"
        )
        return

    username = username.strip()
    password = password.strip()
    context.user_data["awaiting_login"] = False

    status_msg = await update.message.reply_text("⏳ Đang xác thực đăng nhập...")

    result = await login(
        url=url,
        username=username,
        password=password
    )

    if result.get("success"):
        real_base = result.get("base_url") or url
        context.user_data["auth_session"] = {
            "url": real_base,
            "username": username,
            "password": password,
            "hospital_name": hospital_name,
            "display_name": hospital_display,
        }

        # Lưu thông tin khoa phòng vào DB ngầm nếu cần
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
            f"🏥 Bệnh viện: <b>{hospital_display}</b>\n"
            f"👤 Tài khoản: <code>{username}</code>\n"
            f"{info_line}"
            f"📅 <b>Chọn khoảng thời gian xuất danh sách bệnh nhân:</b>",
            parse_mode="HTML",
            reply_markup=keyboard,
        )
        log.info(f"Login OK: {hospital_display} | {username} | user_id={update.effective_user.id}")
    else:
        await status_msg.edit_text(
            f"❌ <b>Đăng nhập thất bại!</b>\n"
            f"🏥 Bệnh viện: <b>{hospital_display}</b>\n"
            f"👤 Tài khoản: <code>{username}</code>\n"
            f"🔴 {result.get('message', 'Không đăng nhập được')}",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏥 Chọn lại bệnh viện", callback_data="show_hospitals")]
            ])
        )
        log.warning(f"Login FAIL: {hospital_display} | {username} | {result.get('message')}")


async def export_date_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Xử lý khi người dùng chọn khoảng thời gian để xuất file."""
    query = update.callback_query
    await query.answer()

    auth_session = context.user_data.get("auth_session")
    if not auth_session:
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
    hosp_name = auth_session.get("hospital_name", "hospital")
    disp_name = auth_session.get("display_name", hosp_name)

    status_msg = await query.message.reply_text(
        f"⏳ <b>Đang thu thập và xuất danh sách bệnh nhân...</b>\n"
        f"🏥 Bệnh viện: <b>{disp_name}</b>\n"
        f"📅 Khoảng thời gian: <b>{label}</b>\n"
        f"<i>Hệ thống đang quét các Stored Procedure tiếp nhận, vui lòng chờ...</i>",
        parse_mode="HTML"
    )

    try:
        exp_res = await export_hospital_patients_excel(
            url=auth_session["url"],
            username=auth_session["username"],
            password=auth_session["password"],
            hospital_name=hosp_name,
            date_option=date_opt,
            fetch_all=True,
        )
    except Exception as e:
        log.exception(f"Lỗi khi xuất file Excel cho {disp_name}: {e}")
        await status_msg.edit_text(
            f"❌ <b>Lỗi trong quá trình xuất file:</b> <code>{e}</code>",
            parse_mode="HTML",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔄 Thử lại", callback_data="show_date_options")],
                [InlineKeyboardButton("🏥 Chọn bệnh viện khác", callback_data="show_hospitals")],
            ])
        )
        return

    buttons = [
        [InlineKeyboardButton("🔄 Xuất khoảng thời gian khác", callback_data="show_date_options")],
        [InlineKeyboardButton("🏥 Đổi bệnh viện khác", callback_data="show_hospitals")],
    ]

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


async def show_date_options_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Hiển thị lại menu chọn ngày cho phiên làm việc hiện tại."""
    query = update.callback_query
    await query.answer()

    auth_session = context.user_data.get("auth_session")
    if not auth_session:
        await query.message.reply_text(
            "⚠️ Phiên làm việc đã hết hạn. Vui lòng chọn lại bệnh viện và đăng nhập lại.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏥 Chọn Bệnh Viện", callback_data="show_hospitals")]
            ]),
            parse_mode="HTML"
        )
        return

    disp_name = auth_session.get("display_name", "Bệnh viện")
    keyboard = InlineKeyboardMarkup(DATE_OPTION_BUTTONS)
    await query.message.reply_text(
        f"📅 <b>Chọn khoảng thời gian xuất danh sách bệnh nhân cho {disp_name}:</b>",
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
