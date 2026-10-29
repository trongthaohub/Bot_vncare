import pathlib

SRC = pathlib.Path(__file__).parent / "data" / "loc.txt"
DST = pathlib.Path(__file__).parent / "data" / "loc_formatted.txt"

lines_out = []
total, skipped = 0, 0
with SRC.open("r", encoding="utf-8", errors="replace") as f:
    header = f.readline()
    for raw in f:
        raw = raw.rstrip("\r\n")
        if not raw.strip():
            skipped += 1
            continue
        parts = raw.split("\t")
        if len(parts) < 2:
            skipped += 1
            continue
        user, pwd = parts[0].strip(), "\t".join(parts[1:]).strip()
        if not user:
            skipped += 1
            continue
        if pwd.startswith("[fail"):
            skipped += 1
            continue
        lines_out.append(f"{user}:{pwd}")
        total += 1

DST.write_text("\n".join(lines_out) + "\n", encoding="utf-8")
print(f"Done: {total} dong -> {DST.name}, bo qua: {skipped}")
