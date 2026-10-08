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


def resolve_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Finds first matching column name ignoring case and whitespace."""
    for candidate in candidates:
        for col in df.columns:
            if str(col).strip().lower() == candidate.lower():
                return col
    return None


def find_default_file(extensions: list[str], preferred_names: list[str] = []) -> Path | None:
    """Auto-detects files matching preferred names or extensions in the script directory."""
    script_dir = Path(__file__).resolve().parent
    for name in preferred_names:
        candidate = script_dir / name
        if candidate.exists():
            return candidate
    for ext in extensions:
        matches = list(script_dir.glob(f"*{ext}"))
        if matches:
            return matches[0]
    return None


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
    min_size: int = 20,
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


def select_file(title: str, filetypes: list[tuple[str, str]]) -> str | None:
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

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Generate certificates from Excel. Supports sports, technical events, and general participants.\n"
            "PDFs are generated and organized into subfolders (e.g. by Sport or College).\n"
            "Point --output at your Google Drive sync folder or local folder."
        )
    )

    parser.add_argument("--excel", help="Path to Excel file (auto-detected if in current folder)")
    parser.add_argument("--template", help="Path to certificate template image (auto-detected if in current folder)")
    parser.add_argument("--sheet", help="Specific sheet name to read from Excel (default: 'Participants' or first sheet)")
    parser.add_argument("--output", default="generated_certificates", help="Output folder (default: generated_certificates)")
    parser.add_argument("--font", help="Path to custom TTF font file")

    # PARTICIPANT NAME PARAMETERS
    parser.add_argument("--x-pct", type=float, default=0.500, help="Name horizontal center fraction (0.500 = center)")
    parser.add_argument("--y-pct", type=float, default=0.5104, help="Name vertical center fraction (centered between pill and 1st dotted line)")
    parser.add_argument("--font-size-px", type=int, default=59, help="Starting font size in pixels for participant name (default: 59px)")
    parser.add_argument("--min-font-size-px", type=int, default=30, help="Minimum font size in pixels when scaling (default: 30px)")
    parser.add_argument("--max-width-pct", type=float, default=0.75, help="Maximum width allowed fraction before auto-shrinking (default: 0.75)")
    parser.add_argument("--color", default="#1a1a1a", help="Text color in hex format (default: #1a1a1a)")

    # Absolute pixel overrides for Name
    parser.add_argument("--name-x", type=int, default=None, help="Absolute X position override in px for name")
    parser.add_argument("--name-y", type=int, default=None, help="Absolute Y position override in px for name")

    # COLLEGE PARAMETERS
    parser.add_argument("--college-y-pct", type=float, default=0.5711, help="College vertical center fraction (centered between the two dotted lines)")
    parser.add_argument("--other-size", type=int, default=38, help="Font size for College name (default: 38px)")

    # SPORT / EVENT PARAMETERS
    parser.add_argument("--sport-x-pct", "--event-x-pct", dest="sport_x_pct", type=float, default=0.3934,
                        help="Sport horizontal center fraction in dotted line blank (default: 0.3934 / ~690px)")
    parser.add_argument("--sport-y-pct", "--event-y-pct", dest="sport_y_pct", type=float, default=0.6358,
                        help="Sport vertical center fraction resting above dotted line (default: 0.6358 / ~789px)")
    parser.add_argument("--sport-size", type=int, default=28, help="Starting font size for Sport/Event (default: 28px)")
    parser.add_argument("--sport-color", default=None, help="Text color for Sport (default: matches --color)")
    parser.add_argument("--sport-bold", action="store_true", default=True, help="Use bold font for Sport (default: True)")
    parser.add_argument("--no-sport-bold", dest="sport_bold", action="store_false", help="Use regular font for Sport")
    parser.add_argument("--clear-sport-dots", "--mask-sport-dots", dest="clear_sport_dots", action="store_true", default=False,
                        help="Mask dots behind sport name with a white box (default: False, text sits above continuous line)")

    # SUBFOLDER GROUPING
    parser.add_argument("--group-by", choices=["auto", "sport", "college", "sport-college", "none"], default="auto",
                        help="Subfolder grouping: 'sport', 'college', 'sport-college', 'none', or 'auto' (default: auto)")

    # DEMO & PREVIEW OPTIONS
    parser.add_argument("--demo", action="store_true", help="Generate 1 sample demo certificate and save preview image")
    parser.add_argument("--demo-index", type=int, default=0, help="Row index (0-based) to use in demo mode (default: 0)")
    parser.add_argument("--limit", type=int, default=None, help="Limit number of certificates to generate")
    parser.add_argument("--preview-image", action="store_true", help="Also save a high-res PNG preview alongside each PDF")

    return parser


# -----------------------------
# Main Logic
# -----------------------------

def main() -> None:
    args = build_parser().parse_args()

    # 1. Resolve Excel file
    excel_path = args.excel
    if not excel_path:
        default_excel = find_default_file([".xlsx", ".xls"], preferred_names=["sports_participants.xlsx"])
        if default_excel:
            excel_path = default_excel
        else:
            excel_path = select_file("Select Excel File", [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])

    # 2. Resolve Template image
    template_path = args.template
    if not template_path:
        default_tpl = find_default_file(
            [".jpg", ".png", ".jpeg"],
            preferred_names=["TechTrove official certificate_page-0001.jpg", "certificate.jpg", "calibration_certificate.jpg"]
        )
        if default_tpl:
            template_path = default_tpl
        else:
            template_path = select_file("Select Certificate Template", [("Image files", "*.png *.jpg *.jpeg"), ("All files", "*.*")])

    if not excel_path or not template_path:
        raise ValueError("Excel file and Template image are required.")

    excel_path = Path(excel_path)
    template_path = Path(template_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 3. Read Excel data
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_to_use = args.sheet if args.sheet else ("Participants" if "Participants" in xl.sheet_names else 0)
        df = pd.read_excel(excel_path, sheet_name=sheet_to_use)
    except Exception:
        df = pd.read_excel(excel_path)

    # Resolve columns dynamically
    name_col = resolve_column(df, ["Participant Name", "Name", "Student Name", "Candidate Name", "Full Name", "Participant"])
    college_col = resolve_column(df, ["College Name", "College", "Institution", "College / Institute", "School / College", "Department"])
    sport_col = resolve_column(df, ["Sport", "Event", "Game", "Sports", "Event Name", "Competition", "Activity"])

    if not name_col:
        raise ValueError(
            f"Excel must contain a participant name column (e.g. 'Participant Name' or 'Name'). Found columns: {list(df.columns)}"
        )

    template = Image.open(template_path).convert("RGB")
    image_width, image_height = template.size

    # Calculate absolute positions
    center_x = args.name_x if args.name_x is not None else int(image_width * args.x_pct)
    center_y = args.name_y if args.name_y is not None else int(image_height * args.y_pct)
    max_allowed_width = image_width * args.max_width_pct

    college_y = int(image_height * args.college_y_pct)
    sport_center_x = int(image_width * args.sport_x_pct)
    sport_center_y = int(image_height * args.sport_y_pct)
    sport_color = args.sport_color or args.color
    max_sport_width = image_width * 0.22  # Maximum width allowed in sport dotted slot (~380px)

    # Grouping strategy
    group_by = args.group_by
    if group_by == "auto":
        if sport_col:
            group_by = "sport"
        elif college_col:
            group_by = "college"
        else:
            group_by = "none"

    # Handle demo mode or limit
    if args.demo:
        demo_idx = max(0, min(args.demo_index, len(df) - 1))
        df_to_process = df.iloc[[demo_idx]].copy()
        print(f"\n[DEMO MODE] Processing single participant at index {demo_idx}")
    elif args.limit is not None and args.limit > 0:
        df_to_process = df.head(args.limit).copy()
        print(f"\n[LIMIT MODE] Processing first {len(df_to_process)} participants")
    else:
        df_to_process = df.copy()

    print(f"[+] Excel source: {excel_path} ({len(df)} total rows)")
    print(f"[+] Detected columns: Name='{name_col}', College='{college_col}', Sport/Event='{sport_col}'")
    print(f"[+] Certificate resolution: {image_width}x{image_height}px")
    print(f"[+] Grouping strategy: {group_by}")
    print(f"[+] Sport position: ({sport_center_x}px, {sport_center_y}px) [xPct: {args.sport_x_pct:.4f}, yPct: {args.sport_y_pct:.4f}]")
    print(f"[+] Participant name position: ({center_x}px, {center_y}px) [size: {args.font_size_px}px]")
    print(f"[+] College position: y={college_y}px [size: {args.other_size}px]")

    generated = 0

    for index, row in df_to_process.iterrows():
        name = get_cell_text(row[name_col])
        college = get_cell_text(row[college_col]) if college_col else ""
        sport = get_cell_text(row[sport_col]) if sport_col else ""

        if not name:
            continue

        certificate = template.copy()
        draw = ImageDraw.Draw(certificate)

        # 1. Render Participant Name (bold serif, auto-shrinks if long)
        name_font = get_fitted_font(
            text=name,
            font_path=args.font,
            start_size=args.font_size_px,
            min_size=args.min_font_size_px,
            max_allowed_width=max_allowed_width,
            bold=True,
        )
        draw_centered_text(
            draw=draw,
            center_x=center_x,
            center_y=center_y,
            text=name,
            font=name_font,
            fill=args.color,
        )

        # 2. Render College (regular serif, auto-shrinks to never overflow)
        if college:
            college_font = get_fitted_font(
                text=college,
                font_path=args.font,
                start_size=args.other_size,
                min_size=20,
                max_allowed_width=max_allowed_width,
                bold=False,
            )
            draw_centered_text(
                draw=draw,
                center_x=center_x,
                center_y=college_y,
                text=college,
                font=college_font,
                fill=args.color,
            )

        # 3. Render Sport / Event (in designated dotted blank with clean masking)
        if sport:
            sport_font = get_fitted_font(
                text=sport,
                font_path=args.font,
                start_size=args.sport_size,
                min_size=16,
                max_allowed_width=max_sport_width,
                bold=args.sport_bold,
            )

            if args.clear_sport_dots:
                bbox_s = draw.textbbox((0, 0), sport, font=sport_font)
                sw = bbox_s[2] - bbox_s[0]
                sh = bbox_s[3] - bbox_s[1]
                pad_x = 6
                pad_y = 2
                draw.rectangle(
                    [
                        sport_center_x - sw // 2 - pad_x,
                        sport_center_y - sh // 2 - pad_y,
                        sport_center_x + sw // 2 + pad_x,
                        sport_center_y + sh // 2 + pad_y,
                    ],
                    fill="#ffffff",
                )

            draw_centered_text(
                draw=draw,
                center_x=sport_center_x,
                center_y=sport_center_y,
                text=sport,
                font=sport_font,
                fill=sport_color,
            )

        # 4. Determine destination directory based on grouping strategy
        if group_by == "sport" and sport:
            dest_dir = output_dir / sanitize_filename(sport)
        elif group_by == "college" and college:
            dest_dir = output_dir / sanitize_filename(college)
        elif group_by == "sport-college" and (sport or college):
            dest_dir = output_dir / sanitize_filename(sport or "General") / sanitize_filename(college or "General")
        else:
            dest_dir = output_dir

        dest_dir.mkdir(parents=True, exist_ok=True)

        safe_name = sanitize_filename(name)
        pdf_path = dest_dir / f"{safe_name}.pdf"
        if pdf_path.exists():
            pdf_path = dest_dir / f"{safe_name}_{index+1}.pdf"

        # Save PDF
        certificate.save(pdf_path, "PDF", resolution=100.0)

        # In demo mode or if --preview-image is set, save PNG preview
        png_path = None
        if args.demo or args.preview_image:
            png_path = dest_dir / f"{safe_name}.png"
            if png_path.exists():
                png_path = dest_dir / f"{safe_name}_{index+1}.png"
            certificate.save(png_path, "PNG")

        generated += 1

        if args.demo:
            print(f"\n[DEMO CREATED] Details for participant:")
            print(f"  • Name   : {name}")
            print(f"  • College: {college}")
            print(f"  • Sport  : {sport}")
            print(f"  • PDF    : {pdf_path.resolve()}")
            if png_path:
                print(f"  • Image  : {png_path.resolve()}")

    print(f"\n[SUCCESS] Successfully generated {generated} certificate(s)")
    print(f"[OUTPUT] Saved in: {output_dir.resolve()}")


if __name__ == "__main__":
    main()