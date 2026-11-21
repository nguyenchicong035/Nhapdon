"""Read ruled tables directly, or reconstruct unruled and scanned tables by position."""
import csv
import io
import re
import subprocess

import pdfplumber

from .common import LOOKUP, clean, headers, is_header, norm, records


def ocr_words(page, languages):
    image = page.to_image(resolution=300).original
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    try:
        proc = subprocess.run(["tesseract", "stdin", "stdout", "-l", languages, "--psm", "6", "tsv"],
                              input=buffer.getvalue(), capture_output=True, timeout=180)
    except FileNotFoundError as exc:
        raise ValueError("Cần cài Tesseract OCR và thêm vào PATH.") from exc
    if proc.returncode:
        raise ValueError(proc.stderr.decode(errors="replace")[-500:])
    scale = page.width / image.width
    return [dict(text=r["text"], x0=int(r["left"])*scale, top=int(r["top"])*scale)
            for r in csv.DictReader(io.StringIO(proc.stdout.decode("utf-8")), delimiter="\t", quoting=csv.QUOTE_NONE)
            if r["text"].strip()]


def lines(words):
    result = []
    for w in sorted(words, key=lambda w: w["top"]):
        if not result or w["top"] - result[-1][0]["top"] > 4:
            result.append([])
        result[-1].append(w)
    return [sorted(line, key=lambda w: w["x0"]) for line in result]


def columns(line):
    found = []
    i = 0
    while i < len(line):
        for length in range(min(6, len(line)-i), 0, -1):
            key = LOOKUP.get(norm(" ".join(w["text"] for w in line[i:i+length])))
            if key:
                found.append((line[i]["x0"], key))
                i += length
                break
        else:
            i += 1
    return found


def positioned(words, layout, page_no):
    rows = []
    pending = {}
    for line in lines(words):
        detected = columns(line)
        if is_header(dict((key, x) for x, key in detected)):
            layout = detected
            continue
        text = " ".join(w["text"] for w in line)
        if not layout or re.match(r"^(Trang|Page)\s+\d+", text, re.I):
            continue
        cells = {}
        for word in line:
            key = next((k for x, k in reversed(layout) if word["x0"] >= x-5), None)
            if key:
                cells.setdefault(key, []).append(word["text"])
        part = {k: " ".join(v) for k, v in cells.items()}
        if part.get("qty") and pending.get("qty"):
            rows.append(pending)
            pending = {}
        for k, v in part.items():
            pending[k] = (pending.get(k, "") + " " + v).strip()
    if pending:
        rows.append(pending)
    if not layout:
        raise ValueError(f"Trang {page_no}: không nhận diện được tiêu đề bảng.")
    keys = [k for _, k in layout]
    titles = [next(label for label, key in LOOKUP.items() if key == k) for k in keys]
    return list(records([titles] + [[r.get(k, "") for k in keys] for r in rows], f"Trang {page_no}")), layout


def parse(path, languages):
    from .platform_pdf import read
    specific = read(path)
    if specific is not None:
        return specific
    output, layout, grid_header = [], [], None
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_rows = []
            for table in page.extract_tables():
                if any(is_header(headers(row)) for row in table):
                    grid_header = next(row for row in table if is_header(headers(row)))
                elif grid_header and table and len(table[0]) == len(grid_header):
                    table = [grid_header, *table]
                else:
                    continue
                page_rows.extend(records(table, f"Trang {page.page_number}"))
            if not page_rows:
                words = page.extract_words() or ocr_words(page, languages)
                page_rows, layout = positioned(words, layout, page.page_number)
            if not page_rows:
                raise ValueError(f"Trang {page.page_number}: không đọc được dòng hàng.")
            for item in page_rows:
                for key in ("sku", "order"):
                    item[key] = clean(item.get(key)).replace("\n", "")
                if not item["order"] and output:
                    item["order"] = output[-1]["order"]
                output.append(item)
    return output
