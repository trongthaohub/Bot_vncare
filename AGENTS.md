# AGENTS.md

## Project
- Async `python-telegram-bot` application for logging into VnCare HIS hospital systems and exporting patient lists to Excel. Keep code, comments, and user-facing text in Vietnamese.
- `bot.py` is the Telegram entrypoint and contains handlers, account parsing, filtering, export orchestration, and `main()`.
- `libs/login.py` owns VnCare authentication; `Controller/bvlaichau.py` is the controller imported by `bot.py`. `Controller/bvbinhthuan.py` is standalone and not wired in.
- `database/` creates/uses SQLite `database/bot.db` at import time. `libs/sel_dept.py` caches departments/rooms there.
- `hospitals.json` is the hospital catalog. `tinh_tp.json` and `libs/tinh_tp.py` drive the 63-province filter, including merged-province matching.

## Commands and configuration
- Install in the ignored virtual environment: `pip install -r requirements.txt`; run with `python bot.py`.
- Runtime configuration is `config.py`; the README and `.env.example` describe an older setup. Never print, replace, or commit the real `BOT_TOKEN` in `config.py`.
- There is no automated test/lint/formatter configuration. Minimum verification after Python edits: `python -m py_compile bot.py Controller/bvlaichau.py libs/tinh_tp.py`.
- `test/export_noi.py` is a manual live-backend script with hardcoded credentials; do not run it as a normal test.

## Behavior that is easy to break
- `login()` returns `error_kind` (`ok`, `auth`, `network`, or `unexpected`); use `is_retryable()` and do not treat network failures as bad credentials.
- Normal export input must include a VnCare URL per account (`url:user:pass`, `url user pass`, full `/vnpthis` URL, or `url|user|pass`). `normalize_vncare_url()` is the security gate and only permits `*.vncare.vn` hosts.
- The filter flow accepts `user:pass`, asks for province and count, samples up to five patients across candidate systems, and returns `url:user:pass`. Large scans show progress/ETA and can be stopped; keep `Application.builder().concurrent_updates(True)` so stop callbacks are processed.
- `httpx` calls intentionally use `trust_env=False` in the relevant clients; removing it can make local proxy/TLS settings break VnCare requests.
- Telegram messages must stay under 4096 characters; use `safe_edit_or_reply()` for long output. File export uploads use long timeouts and may be slow.

## Operations and repository hygiene
- Production is the `botvncare` systemd service on `/root/Bot_vncare`; use `systemctl stop/restart botvncare`, not `kill`, because restart-on-failure respawns it. Inspect app output in `/root/Bot_vncare/logs/bot.log`; `journalctl -u botvncare` mainly shows service lifecycle.
- Telegram `409 Conflict` means another poller is using the same token. `main()` also binds UDP `127.0.0.1:47615` as a single-instance guard.
- `logs/bot.log`, `database/bot.db`, and generated `__pycache__/` files may appear as noisy changes. Never use `git add -A`; stage only files intentionally changed.
