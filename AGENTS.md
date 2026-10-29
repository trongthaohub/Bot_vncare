# AGENTS.md

## What this is
Telegram bot (`python-telegram-bot` v20+, async) that logs into VnCare HIS sites for many hospitals, lists departments/rooms, and exports patient lists to Excel. Code, comments and user-facing strings are Vietnamese — keep them that way.

## Run / config
- Install: `pip install -r requirements.txt` (inside `venv/`, which is gitignored).
- Run: `python bot.py`.
- Config comes from **`config.py`**, not `.env`. `README.md` and `.env.example` are stale — ignore them.
- `config.py` is tracked in git with the real `BOT_TOKEN` despite being listed in `.gitignore`. Never print/echo it, and never commit token changes.
- Logs: `logs/bot.log` (WatchedFileHandler) + colored console output.
- No test suite, linter or formatter config. `test/export_noi.py` is a manual script that hits the real VnCare backend with hardcoded credentials. Safe check for edits: `python -c "import ast,io;ast.parse(open('bot.py',encoding='utf-8').read())"`.

## Layout
- `bot.py` — entire Telegram layer: handlers, account-line parsing, export orchestration, `main()`. Single file, ~1200 lines.
- `database/` — SQLite `database/bot.db`, schema auto-created; module-level singleton `db = Database()` at import time.
- `libs/login.py` — `login(url, username, password)` against `{base}/vnpthis/servlet/login.ValidateUser`. Callers **must** read `result["error_kind"]` (`"ok" | "auth" | "network" | "unexpected"`) and use `is_retryable()`; treating every failure as bad credentials wrongly rejects valid accounts on flaky networks.
- `libs/sel_dept.py` — departments/rooms, cached in DB, `force_refresh` to bypass cache.
- `libs/get_menu.py`, `libs/export_excel.py` (openpyxl helpers, `DEFAULT_FIELDS`).
- `Controller/bvlaichau.py` — hospital-specific export flow (uuid, procedure call, patient rows); the only controller wired into `bot.py`.
- `Controller/bvbinhthuan.py` — near-duplicate variant, **not** imported by `bot.py`.
- Controllers prepend the project root to `sys.path` so they can also run standalone.
- `hospitals.json` — 35-hospital catalog `{name, domain, link}`; UI matches keys case-insensitively against name/domain. `raw.json` is scratch source data.
- `tinh_tp.json` — 63 provinces (pre-2025 list) driving the filter UI. `libs/tinh_tp.py` maps province → hospital (`province_hospitals`, heuristic name match + always `vnpt-his.vncare.vn` and `bvdakhoa.vncare.vn`) and matches `DIACHI` (`address_matches_province`).
- Filter flow (`🔍 Lọc Bệnh Viện`): file/`.txt` paste `user:pass` → pick Tỉnh/Thành (`prov:<code>`) → enter count → `process_filter_file` logs into each candidate system and keeps accounts whose 1–5 sampled patients have `DIACHI` in that province; one `.txt` of `url:user:pass` is sent at the end.

## Gotchas
- Telegram `409 Conflict: terminated by other getUpdates request` = two pollers on the same token. Production runs as systemd unit `botvncare` (`enabled`, `Restart=on-failure`), so `kill <pid>` just respawns it — use `systemctl stop botvncare`. Logs: `journalctl -u botvncare` (lifecycle only) and `/root/Bot_vncare/logs/bot.log`. `main()` also binds UDP `127.0.0.1:47615` as a single-instance guard.
- Telegram caps messages at 4096 chars. Route long output through `safe_edit_or_reply()` (chunks at 3800); direct `edit_message_text`/`reply_text` with unbounded text causes `400 Text is too long`.
- `httpx` clients to VnCare deliberately use `trust_env=False` in places to bypass local MITM proxies (Reqable/9router) that break TLS. Don't remove.
- Account input requires the hospital URL on every line (`bvabc.vncare.vn:user:pass`, `bvabc.vncare.vn user pass`, full `https://…/vnpthis:user:pass`, or `url|user|pass`); plain `user:pass` is only accepted in the account-filter flow where the hospital is already chosen. A one-line paste separated by commas/semicolons is split before each vncare URL. Lines are stripped of surrounding `,;` because stray commas silently corrupt usernames. `normalize_vncare_url()` is the single gate — only `*.vncare.vn` hosts are allowed, everything else is dropped.
- `data/loc.txt` is TSV with header `Email/Username\tPassword` and `[fail` sentinel passwords. `convert_loc.py` emits `data/loc_formatted.txt` as `user:pass`, dropping fails.
- Git state is noisy: `__pycache__/`, `logs/bot.log` and `database/bot.db` are tracked (deletions/modifications show up). Never `git add -A`; stage only the files you touched.
- `.watchfiles.toml` excludes logs/pycache/db if you use watchfiles for auto-reload.
