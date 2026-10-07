import argparse
import re
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


# -----------------------------
# Utility Functions
# -----------------------------

def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:150] if cleaned else "certificate"


def get_cell_text(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def load_font(font_path: str | None, size: int, bold: bool = True) -> ImageFont.ImageFont:
    if font_path and Path(font_path).exists():
        try:
            return ImageFont.truetype(font_path, size)
        except OSError:
            pass

    # Preferred font stack: Georgia (bold), Times New Roman (bold), Serif fallbacks
    font_candidates = (
        ["georgiab.ttf", "timesbd.ttf", "georgia.ttf", "times.ttf", "arialbd.ttf", "arial.ttf"]
        if bold
        else ["georgia.ttf", "times.ttf", "georgiab.ttf", "timesbd.ttf", "arial.ttf"]
    )

    for font_name in font_candidates:
        try:
            return ImageFont.truetype(font_name, size)
        except OSError:
            continue

    return ImageFont.load_default()


def get_fitted_font(
    text: str,
    font_path: str | None,
    start_size: int = 84,
    min_size: int = 30,
    max_allowed_width: float = 0,
    bold: bool = True,
) -> ImageFont.ImageFont:
    """Auto-shrinks font size down from start_size to min_size if text width exceeds max_allowed_width."""
    size = start_size
    dummy_img = Image.new("RGB", (1, 1))
    dummy_draw = ImageDraw.Draw(dummy_img)

    while size > min_size:
        font = load_font(font_path, size, bold=bold)
        bbox = dummy_draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        if max_allowed_width <= 0 or text_width <= max_allowed_width:
            return font
        size -= 2

    return load_font(font_path, min_size, bold=bold)


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    center_x: int,
    center_y: int,
    text: str,
    font: ImageFont.ImageFont,
    fill: str,
) -> None:
    """Draw text precisely centered at (center_x, center_y)."""
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    x = center_x - bbox[0] - text_width // 2
    y = center_y - bbox[1] - text_height // 2

    draw.text((x, y), text, fill=fill, font=font)


def draw_text_on_line(
    draw: ImageDraw.ImageDraw,
    x: int,
    y: int,
    text: str,
    font: ImageFont.ImageFont,
    fill: str,
) -> None:
    bbox = draw.textbbox((0, 0), text, font=font)
    y_adjusted = y - bbox[1] - (bbox[3] - bbox[1]) // 2

    draw.text((x, y_adjusted), text, fill=fill, font=font)


def select_file(title, filetypes):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        file = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return file or None
    except Exception:
        return None


# -----------------------------
# Argument Parser
# -----------------------------

def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Generate certificates from Excel. PDFs are saved one per row, "
            "organized into subfolders by College name when a 'College' column exists. "
            "Point --output at your Google Drive sync folder (e.g. C:\\Users\\<you>\\My Drive\\Certificates) "
            "and files appear in Drive automatically."
        )
    )

    parser.add_argument("--excel", help="Path to Excel file")
    parser.add_argument("--template", help="Path to certificate template image")
    parser.add_argument("--output", default="generated_certificates", help="Output folder")
    parser.add_argument("--font", help="Path to TTF font file")

    # PARTICIPANT NAME PARAMETERS
    parser.add_argument("--x-pct", type=float, default=0.500, help="Horizontal position fraction (0.500 = center)")
    parser.add_argument("--y-pct", type=float, default=0.5104, help="Vertical position fraction (name centered between pill and first dotted line)")
    parser.add_argument("--font-size-px", type=int, default=59, help="Starting font size in pixels (default: 59px)")
    parser.add_argument("--min-font-size-px", type=int, default=30, help="Minimum font size in pixels when scaling (default: 30px)")
    parser.add_argument("--max-width-pct", type=float, default=0.75, help="Maximum width allowed fraction before auto-shrinking (default: 0.75)")
    parser.add_argument("--color", default="#1a1a1a", help="Text color in hex format (default: #1a1a1a)")

    # Legacy absolute pixel overrides
    parser.add_argument("--name-x", type=int, default=None, help="Absolute X position override in px")
    parser.add_argument("--name-y", type=int, default=None, help="Absolute Y position override in px")

    # Other field positions
    parser.add_argument("--college-y-pct", type=float, default=0.5711, help="College vertical position fraction (college centered between the two dotted lines)")
    parser.add_argument("--event-y-pct", type=float, default=0.660, help="Event vertical position fraction")
    parser.add_argument("--other-size", type=int, default=38, help="Font size for college/event")

    return parser


# -----------------------------
# Main Logic
# -----------------------------

def main():
    args = build_parser().parse_args()

    excel_path = args.excel or select_file(
        "Select Excel File",
        [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )

    template_path = args.template or select_file(
        "Select Certificate Template",
        [("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")]
    )

    if not excel_path or not template_path:
        raise ValueError("Excel file and Template image are required.")

    excel_path = Path(excel_path)
    template_path = Path(template_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel not found: {excel_path}")

    if template_path and not template_path.exists():
        for candidate in [template_path.with_name(template_path.name + ".jpg"), Path("certificate.jpg.jpg"), Path("calibration_certificate.jpg")]:
            if candidate.exists():
                template_path = candidate
                break

    if not template_path or not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(excel_path)

    if "Name" not in df.columns:
        raise ValueError("Excel must contain column: Name")

    template = Image.open(template_path).convert("RGB")
    image_width, image_height = template.size

    # Calculate absolute positions from parameters
    center_x = args.name_x if args.name_x is not None else int(image_width * args.x_pct)
    center_y = args.name_y if args.name_y is not None else int(image_height * args.y_pct)
    max_allowed_width = image_width * args.max_width_pct

    other_font = load_font(args.font, args.other_size, bold=False)

    generated = 0

    print(f"[+] Certificate resolution: {image_width}x{image_height}px")
    print(f"[+] Name center position: ({center_x}px, {center_y}px) [xPct: {args.x_pct}, yPct: {args.y_pct}]")
    print(f"[+] Font config: Georgia/Times New Roman (Bold), Start size: {args.font_size_px}px, Min size: {args.min_font_size_px}px, Max width: {int(max_allowed_width)}px ({args.max_width_pct*100}%), Color: {args.color}")

    for index, row in df.iterrows():
        name = get_cell_text(row["Name"])
        college = get_cell_text(row["College"]) if "College" in df.columns else ""
        event = get_cell_text(row["Event"]) if "Event" in df.columns else ""

        if not name:
            continue

        certificate = template.copy()
        draw = ImageDraw.Draw(certificate)

        # Auto-fitted bold serif font for Participant Name
        name_font = get_fitted_font(
            text=name,
            font_path=args.font,
            start_size=args.font_size_px,
            min_size=args.min_font_size_px,
            max_allowed_width=max_allowed_width,
            bold=True,
        )

        # Render Participant Name horizontally and vertically centered at (center_x, center_y)
        draw_centered_text(
            draw=draw,
            center_x=center_x,
            center_y=center_y,
            text=name,
            font=name_font,
            fill=args.color,
        )

        # College (if present in Excel)
        if college:
            college_y = int(image_height * args.college_y_pct)
            draw_centered_text(
                draw=draw,
                center_x=center_x,
                center_y=college_y,
                text=college,
                font=other_font,
                fill=args.color,
            )

        # Event (if present in Excel)
        if event:
            event_y = int(image_height * args.event_y_pct)
            draw_centered_text(
                draw=draw,
                center_x=center_x,
                center_y=event_y,
                text=event,
                font=other_font,
                fill=args.color,
            )

        safe_name = sanitize_filename(name)

        college_dir = output_dir
        if college:
            college_dir = output_dir / sanitize_filename(college)
            college_dir.mkdir(parents=True, exist_ok=True)

        pdf_path = college_dir / f"{safe_name}.pdf"

        if pdf_path.exists():
            pdf_path = college_dir / f"{safe_name}_{index+1}.pdf"

        certificate.save(pdf_path, "PDF", resolution=100.0)
        generated += 1

    print(f"\n[SUCCESS] Successfully generated {generated} certificate(s)")
    print(f"[OUTPUT] Saved in: {output_dir.resolve()}")


if __name__ == "__main__":
    main()