"""Generate reproducible synthetic PDF test samples with Cyrillic text, tables, and images.

Zero external dependencies beyond PyMuPDF (fitz) and Pillow.
Zero client data, zero confidential paths.
"""

from __future__ import annotations

import io
from pathlib import Path
import fitz
from PIL import Image, ImageDraw


def create_diagram_image(page_idx: int, width: int = 400, height: int = 250) -> bytes:
    """Generate a clean technical diagram bitmap to embed in test PDFs."""
    img = Image.new("RGB", (width, height), color=(240, 244, 248))
    draw = ImageDraw.Draw(img)
    # Border
    draw.rectangle([(10, 10), (width - 10, height - 10)], outline=(30, 60, 100), width=2)
    # Grid lines
    for x in range(30, width - 20, 40):
        draw.line([(x, 15), (x, height - 15)], fill=(210, 220, 230), width=1)
    for y in range(30, height - 20, 30):
        draw.line([(15, y), (width - 15, y)], fill=(210, 220, 230), width=1)
    # Truss / Structural diagram
    points = [(50, height - 40), (width // 2, 40), (width - 50, height - 40)]
    draw.polygon(points, outline=(200, 50, 50), width=3)
    draw.line([(50, height - 40), (width - 50, height - 40)], fill=(50, 120, 200), width=3)
    draw.line([(width // 2, 40), (width // 2, height - 40)], fill=(50, 150, 50), width=2)
    # Annotation box
    draw.rectangle([(width - 140, 20), (width - 20, 60)], fill=(255, 255, 255), outline=(100, 100, 100))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def generate_sample_pdf(target_path: Path, num_pages: int = 5) -> Path:
    target_path = Path(target_path).resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)

    doc = fitz.open()
    for p in range(num_pages):
        page_num = p + 1
        page = doc.new_page(width=595, height=842)  # A4

        # Header with Cyrillic text
        page.insert_text(
            fitz.Point(40, 45),
            f"F-ENGINEERING · СИНТЕТИЧЕСКИЙ ОБРАЗЕЦ S03 / СТР. {page_num} ИЗ {num_pages}",
            fontsize=12,
            fontname="helv",
        )
        page.draw_line(fitz.Point(40, 52), fitz.Point(555, 52), color=(0.2, 0.4, 0.6), width=1)

        # Embedded technical diagram image
        img_data = create_diagram_image(page_num)
        page.insert_image(fitz.Rect(40, 65, 340, 240), stream=img_data)

        # Right-side summary box
        page.draw_rect(fitz.Rect(355, 65, 555, 240), color=(0.3, 0.3, 0.3), width=1, fill=(0.95, 0.97, 1.0))
        page.insert_text(fitz.Point(365, 85), "ХАРАКТЕРИСТИКИ УЗЛА:", fontsize=10)
        page.insert_text(fitz.Point(365, 110), f"• Схема: Ферма Ф-{page_num:02d}", fontsize=9)
        page.insert_text(fitz.Point(365, 130), "• Сталь: С255 ГОСТ 27772", fontsize=9)
        page.insert_text(fitz.Point(365, 150), f"• Нагрузка: {120 + page_num * 15} кН", fontsize=9)
        page.insert_text(fitz.Point(365, 170), "• Прогиб: L/450 (допустимый)", fontsize=9)
        page.insert_text(fitz.Point(365, 190), "• Сварные швы: ГОСТ 5264-80", fontsize=9)
        page.insert_text(fitz.Point(365, 215), "• Контроль: УЗК 100%", fontsize=9)

        # Table header
        table_top = 260
        page.draw_rect(fitz.Rect(40, table_top, 555, table_top + 22), color=(0.2, 0.3, 0.4), fill=(0.85, 0.90, 0.95))
        page.insert_text(fitz.Point(45, table_top + 15), "Поз.", fontsize=9)
        page.insert_text(fitz.Point(75, table_top + 15), "Наименование элемента и ГОСТ", fontsize=9)
        page.insert_text(fitz.Point(320, table_top + 15), "Сечение", fontsize=9)
        page.insert_text(fitz.Point(410, table_top + 15), "Длина, мм", fontsize=9)
        page.insert_text(fitz.Point(480, table_top + 15), "Масса, кг", fontsize=9)

        # Table rows
        for r in range(22):
            row_y = table_top + 22 + r * 22
            bg_col = (0.97, 0.97, 0.97) if r % 2 == 1 else (1.0, 1.0, 1.0)
            page.draw_rect(fitz.Rect(40, row_y, 555, row_y + 22), color=(0.8, 0.8, 0.8), fill=bg_col)
            item_no = r + 1 + (p * 22)
            page.insert_text(fitz.Point(45, row_y + 15), f"{item_no:02d}", fontsize=8)
            page.insert_text(fitz.Point(75, row_y + 15), f"Балка прокатная двутавровая ГОСТ Р 57837-2017", fontsize=8)
            page.insert_text(fitz.Point(320, row_y + 15), f"25Б1 (С255)", fontsize=8)
            page.insert_text(fitz.Point(410, row_y + 15), f"{3000 + (r * 150)}", fontsize=8)
            page.insert_text(fitz.Point(480, row_y + 15), f"{(r + 1) * 38.4:.1f}", fontsize=8)

        # Footer
        page.draw_line(fitz.Point(40, 785), fitz.Point(555, 785), color=(0.5, 0.5, 0.5), width=0.5)
        page.insert_text(
            fitz.Point(40, 800),
            "Чертежи и спецификации разработаны ООО «Ф-Инжиниринг». Копирование без согласования запрещено.",
            fontsize=7,
        )
        page.insert_text(fitz.Point(490, 800), f"Лист {page_num}", fontsize=8)

    doc.save(str(target_path))
    doc.close()
    return target_path


def ensure_all_samples(samples_dir: Path) -> dict[str, Path]:
    samples_dir.mkdir(parents=True, exist_ok=True)
    samples = {
        "sample_1p.pdf": generate_sample_pdf(samples_dir / "sample_1p.pdf", num_pages=1),
        "sample_5p.pdf": generate_sample_pdf(samples_dir / "sample_5p.pdf", num_pages=5),
        "sample_10p.pdf": generate_sample_pdf(samples_dir / "sample_10p.pdf", num_pages=10),
        "sample_30p.pdf": generate_sample_pdf(samples_dir / "sample_30p.pdf", num_pages=30),
    }
    return samples


if __name__ == "__main__":
    s_dir = Path(__file__).resolve().parent / "samples"
    created = ensure_all_samples(s_dir)
    for name, p in created.items():
        print(f"Created {name}: {p.stat().st_size / 1024:.1f} KB")
