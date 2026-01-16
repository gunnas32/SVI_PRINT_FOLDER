import os
import time
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import streamlit as st

import tkinter as tk
from tkinter import filedialog

import win32print
import win32ui
import win32con

from PIL import Image, ImageDraw, ImageFont
import pypdfium2 as pdfium


SUPPORTED_EXTENSIONS = (".pdf", ".txt", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")

# Win32 DMPAPER_* codes
DMPAPER: Dict[str, int] = {
    "A3": 8,
    "A4": 9,
    "A5": 11,
    "LETTER": 1,
    "LEGAL": 5,
}

# Physical sizes for preview (mm)
PAPER_MM: Dict[str, Tuple[int, int]] = {
    "A3": (297, 420),
    "A4": (210, 297),
    "A5": (148, 210),
    "LETTER": (216, 279),
    "LEGAL": (216, 356),
}

SCALE_MODES = [
    "Fit (no crop)",
    "Fill (crop)",
    "Fit width (may crop)",
    "Fit height (may crop)",
]


@dataclass
class PrintResult:
    ok: bool
    message: str


# -----------------------------
# Folder & file helpers
# -----------------------------
def browse_for_folder() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askdirectory(title="Select a folder to print")
    root.destroy()
    return path or ""


def list_printers() -> List[str]:
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags)
    return [p[2] for p in printers]


def iter_printable_files(folder_path: str) -> List[str]:
    files: List[str] = []
    for name in os.listdir(folder_path):
        full = os.path.join(folder_path, name)
        if os.path.isfile(full) and name.lower().endswith(SUPPORTED_EXTENSIONS):
            files.append(full)
    files.sort(key=lambda p: os.path.basename(p).lower())
    return files


# -----------------------------
# Rendering helpers
# -----------------------------
def _render_pdf_page_to_pil(pdf_path: str, page_index: int, target_px: int = 1600) -> Image.Image:
    pdf = pdfium.PdfDocument(pdf_path)
    page = pdf[page_index]
    w, h = page.get_size()
    scale = target_px / max(w, h)
    bitmap = page.render(scale=scale)
    pil = bitmap.to_pil()
    page.close()
    pdf.close()
    return pil


def _txt_to_pil(text_path: str, max_width: int = 1800) -> Image.Image:
    with open(text_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    lines = content.splitlines() or [""]

    try:
        font = ImageFont.truetype("consola.ttf", 28)
    except Exception:
        font = ImageFont.load_default()

    dummy = Image.new("RGB", (10, 10), "white")
    draw = ImageDraw.Draw(dummy)

    line_heights = []
    max_line_w = 0
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        max_line_w = max(max_line_w, w)
        line_heights.append(max(h, 30))

    padding = 60
    img_w = min(max_width, max_line_w + padding * 2)
    img_h = sum(line_heights) + padding * 2

    img = Image.new("RGB", (img_w, img_h), "white")
    draw = ImageDraw.Draw(img)

    y = padding
    for line, lh in zip(lines, line_heights):
        draw.text((padding, y), line, fill="black", font=font)
        y += lh

    return img


def _load_any_file_as_pil(filepath: str, pdf_page: int = 0, preview_px: int = 1600) -> Tuple[Image.Image, str]:
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".pdf":
        img = _render_pdf_page_to_pil(filepath, pdf_page, target_px=preview_px)
        return img, f"PDF page {pdf_page + 1}"

    if ext == ".txt":
        img = _txt_to_pil(filepath)
        return img, "Text rendered"

    # image
    img = Image.open(filepath)
    return img, "Image"


# -----------------------------
# Scaling / placement logic
# -----------------------------
def _compute_draw_rect(
    src_w: int,
    src_h: int,
    dst_w: int,
    dst_h: int,
    scale_mode: str,
) -> Tuple[int, int, int, int]:
    """
    Returns (draw_w, draw_h, x, y) where x,y is top-left position in dst.
    draw_w/h can exceed dst for crop modes.
    """
    if scale_mode == "Fill (crop)":
        scale = max(dst_w / src_w, dst_h / src_h)
    elif scale_mode == "Fit width (may crop)":
        scale = dst_w / src_w
    elif scale_mode == "Fit height (may crop)":
        scale = dst_h / src_h
    else:  # "Fit (no crop)"
        scale = min(dst_w / src_w, dst_h / src_h)

    draw_w = int(src_w * scale)
    draw_h = int(src_h * scale)
    x = (dst_w - draw_w) // 2
    y = (dst_h - draw_h) // 2
    return draw_w, draw_h, x, y


def _visible_area(draw_w: int, draw_h: int, x: int, y: int, dst_w: int, dst_h: int) -> int:
    """
    Visible area on destination after cropping.
    """
    left = max(0, x)
    top = max(0, y)
    right = min(dst_w, x + draw_w)
    bottom = min(dst_h, y + draw_h)
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def _maybe_rotate_to_best_fit(
    img: Image.Image,
    dst_w: int,
    dst_h: int,
    scale_mode: str,
    auto_rotate: bool,
) -> Tuple[Image.Image, bool]:
    if not auto_rotate:
        return img, False

    src_w, src_h = img.size

    # Not rotated
    dw0, dh0, x0, y0 = _compute_draw_rect(src_w, src_h, dst_w, dst_h, scale_mode)
    area0 = _visible_area(dw0, dh0, x0, y0, dst_w, dst_h)

    # Rotated 90 degrees
    dw1, dh1, x1, y1 = _compute_draw_rect(src_h, src_w, dst_w, dst_h, scale_mode)
    area1 = _visible_area(dw1, dh1, x1, y1, dst_w, dst_h)

    if area1 > area0:
        return img.transpose(Image.Transpose.ROTATE_90), True

    return img, False


def _paste_with_crop(canvas: Image.Image, img: Image.Image, x: int, y: int) -> None:
    """
    Paste img into canvas at (x,y) handling negative offsets and cropping.
    """
    cw, ch = canvas.size
    iw, ih = img.size

    # Intersection in canvas coords
    dst_left = max(0, x)
    dst_top = max(0, y)
    dst_right = min(cw, x + iw)
    dst_bottom = min(ch, y + ih)
    if dst_right <= dst_left or dst_bottom <= dst_top:
        return

    # Corresponding source crop
    src_left = dst_left - x
    src_top = dst_top - y
    src_right = src_left + (dst_right - dst_left)
    src_bottom = src_top + (dst_bottom - dst_top)

    crop = img.crop((src_left, src_top, src_right, src_bottom))
    canvas.paste(crop, (dst_left, dst_top))


def build_page_preview(
    content_img: Image.Image,
    paper_key: str,
    landscape: bool,
    scale_mode: str,
    auto_rotate: bool,
    preview_width_px: int = 900,
) -> Tuple[Image.Image, Dict[str, str]]:
    """
    Creates a visual preview showing the selected paper + placement.
    Returns preview image and debug info.
    """
    w_mm, h_mm = PAPER_MM[paper_key]
    if landscape:
        w_mm, h_mm = h_mm, w_mm

    # Canvas size
    page_w = preview_width_px
    page_h = max(200, int(preview_width_px * (h_mm / w_mm)))
    canvas = Image.new("RGB", (page_w, page_h), "white")
    draw = ImageDraw.Draw(canvas)

    # Border
    draw.rectangle((5, 5, page_w - 6, page_h - 6), outline=(0, 0, 0), width=2)

    # Apply auto rotate (preview)
    img, rotated = _maybe_rotate_to_best_fit(content_img, page_w - 20, page_h - 20, scale_mode, auto_rotate)

    # Scale & place inside margins
    inner_w = page_w - 20
    inner_h = page_h - 20
    src_w, src_h = img.size

    draw_w, draw_h, x, y = _compute_draw_rect(src_w, src_h, inner_w, inner_h, scale_mode)
    x += 10
    y += 10

    resized = img.convert("RGB").resize((max(1, draw_w), max(1, draw_h)), Image.Resampling.LANCZOS)
    _paste_with_crop(canvas, resized, x, y)

    info = {
        "paper": f"{paper_key} {'Landscape' if landscape else 'Portrait'}",
        "mode": scale_mode,
        "auto_rotate": "On" if auto_rotate else "Off",
        "rotated": "Yes" if rotated else "No",
    }
    return canvas, info


# -----------------------------
# Printing (Win32 GDI)
# -----------------------------
def _get_devmode_for_printer(printer_name: str, paper_code: int, landscape: bool):
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]
        if devmode is None:
            raise RuntimeError("Printer returned no DEVMODE (driver issue).")

        devmode.PaperSize = paper_code
        devmode.Orientation = win32con.DMORIENT_LANDSCAPE if landscape else win32con.DMORIENT_PORTRAIT
        devmode.Fields |= win32con.DM_PAPERSIZE
        devmode.Fields |= win32con.DM_ORIENTATION
        return devmode
    finally:
        win32print.ClosePrinter(hprinter)


class ImageWinDIB:
    @staticmethod
    def from_pil(img: Image.Image):
        from PIL import ImageWin
        return ImageWin.Dib(img)


def _print_pil_image(
    printer_name: str,
    doc_name: str,
    img: Image.Image,
    paper_code: int,
    landscape: bool,
    scale_mode: str,
    auto_rotate: bool,
) -> PrintResult:
    """
    Prints a PIL image by drawing it to the printer DC and scaling to selected paper.
    """
    try:
        devmode = _get_devmode_for_printer(printer_name, paper_code, landscape)

        hdc = win32ui.CreateDC()
        try:
            # Create DC with WINSPOOL driver and DEVMODE (avoids ResetDC)
            hdc.CreateDC("WINSPOOL", printer_name, None, devmode)
        except Exception:
            # Fallback if some drivers reject CreateDC args
            hdc.CreatePrinterDC(printer_name)

        hdc.SetMapMode(win32con.MM_TEXT)

        page_w = hdc.GetDeviceCaps(win32con.HORZRES)
        page_h = hdc.GetDeviceCaps(win32con.VERTRES)

        # Auto rotate to best fit printer page area
        img2, rotated = _maybe_rotate_to_best_fit(img, page_w, page_h, scale_mode, auto_rotate)

        img2 = img2.convert("RGB")
        src_w, src_h = img2.size

        draw_w, draw_h, x, y = _compute_draw_rect(src_w, src_h, page_w, page_h, scale_mode)

        # Resize to what we're drawing
        resized = img2.resize((max(1, draw_w), max(1, draw_h)), Image.Resampling.LANCZOS)
        dib = ImageWinDIB.from_pil(resized)

        hdc.StartDoc(doc_name)
        hdc.StartPage()

        # If x/y negative we still draw; GDI will clip to page
        dib.draw(hdc.GetHandleOutput(), (x, y, x + draw_w, y + draw_h))

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        rtxt = " (rotated)" if rotated else ""
        return PrintResult(True, f"Printed{rtxt}.")
    except Exception as e:
        return PrintResult(False, f"Print failed: {e}")


def print_file(
    printer_name: str,
    filepath: str,
    paper_key: str,
    landscape: bool,
    scale_mode: str,
    auto_rotate: bool,
) -> PrintResult:
    ext = os.path.splitext(filepath)[1].lower()
    paper_code = DMPAPER[paper_key]

    if ext == ".pdf":
        # Print each page of the PDF
        try:
            pdf = pdfium.PdfDocument(filepath)
            n_pages = len(pdf)
            pdf.close()
        except Exception as e:
            return PrintResult(False, f"Could not open PDF: {e}")

        for p in range(n_pages):
            try:
                # Higher quality for print than preview
                img = _render_pdf_page_to_pil(filepath, p, target_px=2400)
                res = _print_pil_image(
                    printer_name=printer_name,
                    doc_name=f"{os.path.basename(filepath)} (page {p+1})",
                    img=img,
                    paper_code=paper_code,
                    landscape=landscape,
                    scale_mode=scale_mode,
                    auto_rotate=auto_rotate,
                )
                if not res.ok:
                    return res
            except Exception as e:
                return PrintResult(False, f"PDF page {p+1} failed: {e}")

        return PrintResult(True, f"Printed PDF ({n_pages} page(s)).")

    if ext == ".txt":
        img = _txt_to_pil(filepath)
        return _print_pil_image(
            printer_name, os.path.basename(filepath), img, paper_code, landscape, scale_mode, auto_rotate
        )

    # images
    try:
        img = Image.open(filepath)
        return _print_pil_image(
            printer_name, os.path.basename(filepath), img, paper_code, landscape, scale_mode, auto_rotate
        )
    except Exception as e:
        return PrintResult(False, f"Could not open image: {e}")


# -----------------------------
# Streamlit UI
# -----------------------------
def main() -> None:
    st.set_page_config(page_title="Folder Printer", layout="centered")
    st.title("🖨️ Folder Printer")
    st.caption("Select a folder, choose printer + paper size, preview placement, then print.")

    if os.name != "nt":
        st.error("This app is Windows-only.")
        return

    # Folder picker state
    if "folder_path" not in st.session_state:
        st.session_state.folder_path = ""

    c1, c2 = st.columns([1, 2])
    with c1:
        if st.button("📁 Browse folder…"):
            picked = browse_for_folder()
            if picked:
                st.session_state.folder_path = picked
    with c2:
        st.text_input("Selected folder", value=st.session_state.folder_path, disabled=True)

    folder_path = st.session_state.folder_path
    if not folder_path:
        st.info("Select a folder to begin.")
        return
    if not os.path.isdir(folder_path):
        st.error("Selected folder is not valid anymore. Please browse again.")
        return

    # Printers
    try:
        printers = list_printers()
    except Exception as e:
        st.error(f"Could not list printers: {e}")
        return
    if not printers:
        st.error("No printers found.")
        return

    printer_name = st.selectbox("Printer", printers)

    # Settings
    with st.sidebar:
        st.subheader("Page settings")

        paper_key = st.selectbox("Paper size", list(DMPAPER.keys()), index=0)
        landscape = st.checkbox("Landscape", value=(paper_key == "A3"))

        scale_mode = st.selectbox("Scaling", SCALE_MODES, index=SCALE_MODES.index("Fit (no crop)"))
        st.caption("To ‘spread horizontally’ on A3, try **Fit width** or enable **Auto-rotate**.")

        auto_rotate = st.checkbox("Auto-rotate to best fit", value=True)
        st.caption("Rotates 90° only if it uses more page area.")

        st.markdown("---")
        delay = st.number_input("Delay between files (seconds)", min_value=0.0, max_value=30.0, value=1.0, step=0.5)

        st.markdown("---")
        st.write("Supported file types:")
        st.code(", ".join(SUPPORTED_EXTENSIONS), language="text")

    files = iter_printable_files(folder_path)

    st.markdown("### Files found")
    if not files:
        st.warning("No printable files found in the selected folder.")
        return

    st.dataframe({"Files": [os.path.basename(f) for f in files]}, use_container_width=True, hide_index=True)

    # -----------------------------
    # Preview
    # -----------------------------
    st.markdown("### Preview")
    sel_name = st.selectbox("Select a file to preview", [os.path.basename(f) for f in files])
    sel_path = next(f for f in files if os.path.basename(f) == sel_name)

    pdf_page = 0
    if sel_path.lower().endswith(".pdf"):
        try:
            pdf = pdfium.PdfDocument(sel_path)
            n_pages = len(pdf)
            pdf.close()
            if n_pages > 1:
                pdf_page = st.number_input("PDF page", min_value=1, max_value=n_pages, value=1, step=1) - 1
        except Exception:
            st.warning("Could not read PDF page count for preview.")

    with st.spinner("Rendering preview..."):
        content_img, content_label = _load_any_file_as_pil(sel_path, pdf_page=pdf_page, preview_px=1400)
        preview_img, info = build_page_preview(
            content_img=content_img,
            paper_key=paper_key,
            landscape=landscape,
            scale_mode=scale_mode,
            auto_rotate=auto_rotate,
            preview_width_px=900,
        )

    st.image(preview_img, caption=f"{info['paper']} • {info['mode']} • Auto-rotate: {info['auto_rotate']} • Rotated: {info['rotated']}")

    # -----------------------------
    # Print
    # -----------------------------
    st.markdown("### Print")
    confirm = st.checkbox("I confirm I want to print all these files", value=False)

    if st.button("🚀 Print folder", type="primary"):
        if not confirm:
            st.error("Tick the confirmation checkbox first.")
            return

        progress = st.progress(0)
        log = st.empty()

        ok_count = 0
        fail_count = 0

        for i, fpath in enumerate(files, start=1):
            fname = os.path.basename(fpath)
            res = print_file(
                printer_name=printer_name,
                filepath=fpath,
                paper_key=paper_key,
                landscape=landscape,
                scale_mode=scale_mode,
                auto_rotate=auto_rotate,
            )

            if res.ok:
                ok_count += 1
                log.write(f"✅ {i}/{len(files)} **{fname}** — {res.message}")
            else:
                fail_count += 1
                log.write(f"⚠️ {i}/{len(files)} **{fname}** — {res.message}")

            progress.progress(int(i / len(files) * 100))
            time.sleep(float(delay))

        st.success(f"Done. Success: {ok_count}, Failed: {fail_count}.")


if __name__ == "__main__":
    main()
