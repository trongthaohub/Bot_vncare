# Telegram Bot Base (python-telegram-bot v20+)

Đây là bộ khung (base) chuẩn cho Telegram Bot sử dụng thư viện `python-telegram-bot` phiên bản bất đồng bộ (async) mới nhất.

## Cấu trúc thư mục
- `bot.py`: Mã nguồn chính của bot (xử lý lệnh `/start`, `/help` và echo tin nhắn).
- `requirements.txt`: Các thư viện phụ thuộc (`python-telegram-bot`, `python-dotenv`).
- `.env.example`: File mẫu chứa cấu hình môi trường.

## Hướng dẫn cài đặt và chạy

1. **Cài đặt thư viện:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Cấu hình Token:**
   - Copy file `.env.example` thành `.env`:
     ```bash
     cp .env.example .env
     ```
   - Điền token bot Telegram của bạn vào file `.env`:
     ```env
     TELEGRAM_TOKEN=123456789:ABCdefGhIJKlmNoPQRsTUVwxyZ
     ```

3. **Chạy bot:**
   ```bash
   python bot.py
   ```
