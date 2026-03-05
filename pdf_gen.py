import io
import os
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Tuple, Optional

try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    PdfReader = None
    PdfWriter = None
from PIL import Image

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# Background image native size (pixels)
BG_W_PX = 1448
BG_H_PX = 2048

# A4 size in points
PAGE_W, PAGE_H = A4
SCALE = PAGE_W / BG_W_PX  # also equals PAGE_H / BG_H_PX approximately

def px_to_pt(x_px: float, y_px: float) -> Tuple[float, float]:
    """Convert image pixel coords (origin top-left) to PDF points (origin bottom-left)."""
    x_pt = x_px * SCALE
    y_pt = (BG_H_PX - y_px) * SCALE
    return x_pt, y_pt

def _draw_mark_rect(c: canvas.Canvas, x_px: float, y_px: float, size_px: float = 16, pad_px: float = 2) -> None:
    """Draw a filled black square mark inside a checkbox area.
    x_px, y_px are the checkbox top-left-ish pixel coordinates on the background image.
    """
    x_pt = (x_px + pad_px) * SCALE
    # convert top-left pixel y to bottom-left point y for rect bottom
    bottom_y_px = y_px + pad_px + size_px
    y_pt = (BG_H_PX - bottom_y_px) * SCALE
    w = size_px * SCALE
    h = size_px * SCALE
    c.saveState()
    c.setFillColorRGB(0, 0, 0)
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(x_pt, y_pt, w, h, stroke=0, fill=1)
    c.restoreState()

def _try_register_tc_font() -> str:
    """Prefer a Traditional Chinese font if present in ./fonts to avoid garbled Chinese in PDF."""
    candidates = [
        ("bkai00mp", os.path.join("fonts", "bkai00mp.ttf")),
        ("gkai00mp", os.path.join("fonts", "gkai00mp.ttf"))
    ]

    here = os.path.dirname(__file__)
    for name, path in candidates:
        full_path = path if os.path.isabs(path) else os.path.join(here, path)
        if os.path.isfile(full_path):
            try:
                pdfmetrics.registerFont(TTFont(name, full_path))
                return name
            except Exception:
                pass

    # Fallback: built-in CID font for Traditional Chinese (may not render well in all viewers)
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
        return "MSung-Light"
    except Exception:
        return "Helvetica"


def _wrap_text(text: str, font_name: str, font_size: int, max_width_pt: float) -> List[str]:
    """Simple CJK-friendly wrapping by characters."""
    if not text:
        return []
    lines = []
    buf = ""
    for ch in text:
        if ch == "\n":
            lines.append(buf)
            buf = ""
            continue
        w = pdfmetrics.stringWidth(buf + ch, font_name, font_size)
        if w <= max_width_pt:
            buf += ch
        else:
            if buf:
                lines.append(buf)
            buf = ch
    if buf:
        lines.append(buf)
    return lines

def _to_int_amount(amount_str: str) -> int:
    if not amount_str:
        return 0
    try:
        d = Decimal(str(amount_str))
        return int(d)  # ignore decimals for box digits
    except (InvalidOperation, ValueError):
        return 0

# Digit boxes boundaries detected from the provided image (pixels)
DIGIT_BOX_XS = [568, 661, 753, 846, 938, 1030, 1123, 1215, 1307]  # 8 boxes => 9 boundaries
DIGIT_CENTER_X = [(DIGIT_BOX_XS[i] + DIGIT_BOX_XS[i+1]) / 2 for i in range(8)]
DIGIT_CENTER_Y_PX = 915  # center of the digit boxes row (pixels)

def _image_to_pdf_bytes(image_path: str) -> bytes:
    """Convert a single image to a 1-page A4 PDF, scaled to fit with margins."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w, page_h = A4
    margin = 24  # points

    with Image.open(image_path) as im:
        im = im.convert("RGB")
        iw, ih = im.size
        img_ratio = iw / ih if ih else 1.0

    max_w = page_w - 2 * margin
    max_h = page_h - 2 * margin
    box_ratio = max_w / max_h if max_h else 1.0

    if img_ratio >= box_ratio:
        draw_w = max_w
        draw_h = max_w / img_ratio
    else:
        draw_h = max_h
        draw_w = max_h * img_ratio

    x = (page_w - draw_w) / 2
    y = (page_h - draw_h) / 2

    c.drawImage(ImageReader(image_path), x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
    c.showPage()
    c.save()
    return buf.getvalue()


def _merge_attachments(base_pdf: bytes, attachment_paths: List[str]) -> bytes:
    """Append attachment files (PDFs or images) after the first page."""
    writer = PdfWriter()

    base_reader = PdfReader(io.BytesIO(base_pdf))
    for p in base_reader.pages:
        writer.add_page(p)

    for pth in attachment_paths:
        if not pth:
            continue
        if not os.path.exists(pth):
            continue
        lower = pth.lower()
        try:
            if lower.endswith(".pdf"):
                r = PdfReader(pth)
                for page in r.pages:
                    writer.add_page(page)
            elif lower.endswith((".png", ".jpg", ".jpeg", ".webp")):
                img_pdf = _image_to_pdf_bytes(pth)
                r = PdfReader(io.BytesIO(img_pdf))
                for page in r.pages:
                    writer.add_page(page)
            else:
                continue
        except Exception:
            continue

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

def build_pdf_bytes(record: Dict, bg_image_path: str, attachment_paths: Optional[List[str]] = None) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # Background
    if os.path.exists(bg_image_path):
        c.drawImage(ImageReader(bg_image_path), 0, 0, width=PAGE_W, height=PAGE_H, mask='auto')

    font = _try_register_tc_font()
    c.setFont(font, 11)

    # Header date: split into 年 / 月 / 日 blanks
    form_date = record.get("form_date", "")
    ymd = None
    if form_date:
        s = str(form_date).strip()
        # accept 'YYYY-MM-DD' or 'YYYY/MM/DD'
        try:
            date_obj = datetime.fromisoformat(s.replace("/", "-")).date()
            ymd = (date_obj.year, date_obj.month, date_obj.day)
        except Exception:
            m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
            if m:
                ymd = (int(m.group(1)), int(m.group(2)), int(m.group(3)))

    if ymd:
        y_px = 346  # baseline just above the thick line under 年/月/日
        # right edges for digits: just before the printed 年/月/日 labels
        x_year, y = px_to_pt(1111, y_px)
        x_month, y = px_to_pt(1199, y_px)
        x_day, y = px_to_pt(1290, y_px)
        c.drawRightString(x_year, y, f"{ymd[0]:04d}")
        c.drawRightString(x_month, y, f"{ymd[1]:02d}")
        c.drawRightString(x_day, y, f"{ymd[2]:02d}")
    else:
        # fallback: print raw string on the blank line (left of 年/月/日)
        s = str(form_date).strip()
        if s:
            x, y = px_to_pt(943, 356)
            c.drawString(x, y, s)

    # Plan code
    x, y = px_to_pt(425, 395)
    c.drawString(x, y, str(record.get("plan_code","")))

    # Purpose
    purpose = str(record.get("purpose_desc",""))
    max_w = (1300 - 403) * SCALE
    lines = _wrap_text(purpose, font, 11, max_w)
    start_x, start_y = px_to_pt(403, 480)
    line_h = 14
    for i, line in enumerate(lines[:3]):
        c.drawString(start_x, start_y - i*line_h, line)

    # Payee checkbox marks (互斥三選一：員工 / 借支充抵 / 逕付廠商)
    mode = str(record.get("payment_mode","") or "").strip()
    is_adv = str(record.get("is_advance_offset","")).lower() in ("true","1","yes")
    is_vendor = str(record.get("is_direct_vendor_pay","")).lower() in ("true","1","yes") or (record.get("payee_type","") == "vendor")

    if mode not in ("employee", "advance", "vendor"):
        if is_adv:
            mode = "advance"
        elif is_vendor:
            mode = "vendor"
        else:
            mode = "employee"

    
    # Checkbox calibration: move slightly right and up to better center within boxes.
    if mode == "employee":
        _draw_mark_rect(c, 400, 528, size_px=18, pad_px=4)
    elif mode == "advance":
        _draw_mark_rect(c, 400, 588, size_px=18, pad_px=4)
    elif mode == "vendor":
        _draw_mark_rect(c, 400, 738, size_px=18, pad_px=4)
    c.setFont(font, 11)

    # Employee fieldsds
    x, y = px_to_pt(562, 546)
    c.drawString(x, y, str(record.get("employee_name","")))
    x, y = px_to_pt(923, 546)
    c.drawString(x, y, str(record.get("employee_no","")))

    # Advance offset
    show_adv = (mode == "advance") or (str(record.get("is_advance_offset","")).lower() in ("true","1","yes"))
    if show_adv:
        _draw_mark_rect(c, 400, 588, size_px=18, pad_px=4)
        c.setFont(font, 11)

        x, y = px_to_pt(548, 667)
        c.drawString(x, y, str(int(record.get("advance_amount") or 0)))
        x, y = px_to_pt(760, 667)
        c.drawString(x, y, str(int(record.get("offset_amount") or 0)))
        x, y = px_to_pt(965, 667)
        c.drawString(x, y, str(int(record.get("balance_refund_amount") or 0)))
        x, y = px_to_pt(1178, 667)
        c.drawString(x, y, str(int(record.get("supplement_amount") or 0)))


    # Vendor fields
    x, y = px_to_pt(600, 760)
    c.drawString(x, y, str(record.get("vendor_name","")))
    x, y = px_to_pt(600, 800)
    c.drawString(x, y, str(record.get("vendor_address","")))
    x, y = px_to_pt(600, 835)
    c.drawString(x, y, str(record.get("vendor_payee_name","")))

    # Receipt no
    x, y = px_to_pt(210, 915)
    c.drawString(x, y, str(record.get("receipt_no","")))

    # Amount digits in boxes
    amt_int = _to_int_amount(str(record.get("amount_total","")))
    if 0 <= amt_int <= 99999999:
        digits = f"{amt_int:08d}"
        c.setFont(font, 14)
        for i, dch in enumerate(digits):
            cx_pt, cy_pt = px_to_pt(DIGIT_CENTER_X[i], DIGIT_CENTER_Y_PX)
            c.drawCentredString(cx_pt, cy_pt-5, dch)
        c.setFont(font, 11)

    # Signatures (place under the labels, above the dotted line)
    sig_y_px = 1048
    sig_specs = [
        ("handler_name", 229),
        ("project_manager_name", 484),
        ("dept_manager_name", 776),
        ("accountant_name", 1051),
    ]
    for key, x_px in sig_specs:
        val = str(record.get(key, "")).strip()
        if val:
            x, y = px_to_pt(x_px, sig_y_px)
            c.drawCentredString(x, y, val)

    c.showPage()
    c.save()

    base_pdf = buf.getvalue()
    if attachment_paths:
        return _merge_attachments(base_pdf, attachment_paths)
    return base_pdf
