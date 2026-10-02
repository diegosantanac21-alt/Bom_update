#!/usr/bin/env python3
"""
Oracle EBS Routing Report splitter.

- Reads range reports from Input_PDFs1.
- Creates one searchable PDF per Part Number in Split_Output2.
- Keeps the original displayed page size and orientation.
- Removes the first range item physically from the cover.
- Places routing content immediately below the cover.
- Paginates only at whitespace gaps to avoid cutting rows.
- Removes the next item's hidden/searchable text.
- Adds the Oracle End of Report footer.

Install:
    py -m pip install pymupdf openpyxl

Run:
    py -u .\split_routing_reports.py
"""

import re
import sys
import traceback
from pathlib import Path

try:
    import pymupdf as fitz
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    print(f"Falta una dependencia: {exc}")
    print("Ejecuta: py -m pip install pymupdf openpyxl")
    raise SystemExit(1)

__version__ = "1.1.0"

# ── Historial ───────────────────────────────────────────────────────────────
# 1.1.0  Arreglados los saltos de página y los blancos grandes:
#        · Las coordenadas de texto venían SIN rotar mientras page.rect está
#          rotado (estos reportes son rotation=90). Se mapean con
#          page.rotation_matrix.
#        · Se descarta el texto que Oracle deja FUERA del área visible: inflaba
#          el contenido medido muy por encima del alto real de la página.
#        · Ya no se fuerza una página nueva al terminar cada página del
#          reporte original; se continúa en la misma hoja si hay sitio.
#        · Los huecos en blanco internos mayores que MAX_BLANK_GAP se colapsan
#          en vez de arrastrarse al documento de salida.
# 1.0.0  Versión original.

INPUT_FOLDER_NAME = "Input_PDFs1"
OUTPUT_FOLDER_NAME = "Split_Output1"
SUMMARY_FILE_NAME = "Routing_Split_Summary.xlsx"
LOG_FILE_NAME = "Routing_Split_Log.txt"

TOP_MARGIN = 18.0
BOTTOM_MARGIN = 26.0
COVER_TO_ROUTING_GAP = 10.0
MIN_SEGMENT_HEIGHT = 36.0
SOURCE_PAGE_GAP = 14.0      # separación entre páginas del reporte original
MAX_BLANK_GAP = 48.0        # hueco en blanco máximo que se conserva (0 = no colapsar)
FOOTER_SHIFT_RIGHT = 24.0

ITEM_RE = re.compile(r"\bItem:\s*(CE-[A-Z0-9]+-[A-Z0-9-]+)\b", re.I)
INVALID_NAME_RE = re.compile(r'[<>:"/\\|?*]')


def visible_rects(page):
    """Word rectangles mapped to the page's DISPLAYED coordinate space.

    Two corrections that the rest of the code depends on:
      * get_text() reports coordinates in the UNROTATED space. On a rotated
        page (these reports are rotation=90) those values do not match
        page.rect, so every vertical calculation based on them is wrong.
        page.rotation_matrix maps them to what the reader actually sees.
      * Oracle leaves words positioned OUTSIDE the visible page. Including
        them stretches the measured content far beyond the page height, which
        is what produced the oversized page breaks and blank gaps.
    """
    matrix = page.rotation_matrix
    limit = page.rect
    result = []
    for word in page.get_text("words"):
        if not str(word[4]).strip():
            continue
        rect = fitz.Rect(word[:4]) * matrix
        if rect.y1 < limit.y0 - 1 or rect.y0 > limit.y1 + 1:
            continue
        if rect.x1 < limit.x0 - 1 or rect.x0 > limit.x1 + 1:
            continue
        result.append(rect)
    return result


def safe_name(value):
    return INVALID_NAME_RE.sub("_", value.strip()).rstrip(". ") or "UNKNOWN_ITEM"


def item_markers(page):
    """Return [(part_number, y0), ...] for true Item header lines.

    Coordinates are mapped to the DISPLAYED space: on a rotated page the raw
    values do not line up with page.rect, and the cover redaction that uses
    this y would then cut the wrong band.
    """
    matrix = page.rotation_matrix
    lines = {}
    for word in page.get_text("words", sort=True):
        raw = fitz.Rect(word[:4]) * matrix
        x0, y0, x1, y1 = raw.x0, raw.y0, raw.x1, raw.y1
        text, block, line, number = word[4], word[5], word[6], word[7]
        row = lines.setdefault((block, line), {"words": [], "box": [x0, y0, x1, y1]})
        row["words"].append((number, text))
        box = row["box"]
        box[0], box[1] = min(box[0], x0), min(box[1], y0)
        box[2], box[3] = max(box[2], x1), max(box[3], y1)

    result = []
    for row in lines.values():
        line_text = " ".join(text for _, text in sorted(row["words"]))
        match = ITEM_RE.search(line_text)
        if match:
            result.append((match.group(1).upper().rstrip(".,;:"), float(row["box"][1])))
    result.sort(key=lambda pair: pair[1])

    unique = []
    for item, y0 in result:
        if not unique or item != unique[-1][0] or abs(y0 - unique[-1][1]) > 2:
            unique.append((item, y0))
    return unique


def first_rect(page, text):
    hits = sorted(page.search_for(text), key=lambda rect: (rect.y0, rect.x0))
    if not hits:
        raise RuntimeError(f"No se encontro '{text}'")
    return hits[0]


def original_style(page, sample="Range"):
    target = first_rect(page, sample)
    found = None
    for block in page.get_text("dict").get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if fitz.Rect(span["bbox"]).intersects(target):
                    found = span
                    break
    if not found:
        return "cour", 8.2, (0, 0, 0), target.height

    resource = None
    span_font = str(found.get("font", "")).replace(" ", "").lower()
    for font_info in page.get_fonts(full=True):
        base_font = str(font_info[3]).replace(" ", "").lower() if len(font_info) > 3 else ""
        if span_font and (span_font in base_font or base_font in span_font):
            resource = str(font_info[4])
            break
    if not resource:
        fonts = page.get_fonts(full=True)
        resource = str(fonts[0][4]) if fonts else "cour"

    size = float(found.get("size", 8.2))
    color_value = int(found.get("color", 0))
    color = (
        ((color_value >> 16) & 255) / 255,
        ((color_value >> 8) & 255) / 255,
        (color_value & 255) / 255,
    )
    origin = found.get("origin")
    baseline = float(origin[1]) - target.y0 if origin else target.height
    return resource, size, color, baseline


def exact_redaction(page, rect, pad_x=1.0, pad_y=0.5):
    page.add_redact_annot(
        fitz.Rect(rect.x0 - pad_x, rect.y0 - pad_y, rect.x1 + pad_x, rect.y1 + pad_y),
        fill=(1, 1, 1),
    )


def insert_original_style(page, rect, text, style):
    resource, size, color, baseline = style
    point = (rect.x0, rect.y0 + baseline)
    try:
        page.insert_text(point, text, fontname=resource, fontsize=size, color=color)
    except Exception:
        page.insert_text(point, text, fontname="cour", fontsize=size, color=color)


def value_on_row(page, label):
    label_box = first_rect(page, label)
    tolerance = max(3, label_box.height * 0.65)
    result = None
    for word in page.get_text("words", sort=True):
        rect = fitz.Rect(word[:4])
        text = str(word[4])
        if abs(rect.y0 - label_box.y0) <= tolerance and rect.x0 > label_box.x1:
            if text.upper().startswith("CE-"):
                result = rect if result is None else result | rect
    return result


def build_cover(source_doc, item):
    """Create a searchable cover and physically remove the first routing."""
    cover = fitz.open()
    cover.insert_pdf(source_doc, from_page=0, to_page=0)
    page = cover[0]

    markers = item_markers(page)
    if markers:
        first_item_y = markers[0][1]
        page.add_redact_annot(
            fitz.Rect(page.rect.x0, max(page.rect.y0, first_item_y - 2), page.rect.x1, page.rect.y1),
            fill=(1, 1, 1),
        )

    style = original_style(page, "Range")
    range_rect = first_rect(page, "Range")
    value_x = range_rect.x0
    item_label = first_rect(page, "Item:")
    revision_label = first_rect(page, "Revision:")
    item_position = fitz.Rect(value_x, item_label.y0, page.rect.width - 18, item_label.y1)
    revision_position = fitz.Rect(value_x, revision_label.y0, page.rect.width - 18, revision_label.y1)

    from_value = value_on_row(page, "Items From:")
    items_from_label = first_rect(page, "Items From:")
    category_label = first_rect(page, "Category Set:")
    to_label = next(
        (rect for rect in sorted(page.search_for("To:"), key=lambda rect: rect.y0)
         if items_from_label.y0 < rect.y0 < category_label.y0),
        None,
    )
    if from_value is None or to_label is None:
        cover.close()
        raise RuntimeError("No se localizaron Items From / To")

    to_value = None
    tolerance = max(3, to_label.height * 0.65)
    for word in page.get_text("words", sort=True):
        rect = fitz.Rect(word[:4])
        text = str(word[4])
        if abs(rect.y0 - to_label.y0) <= tolerance and rect.x0 > to_label.x1:
            if text.upper().startswith("CE-"):
                to_value = rect if to_value is None else to_value | rect
    if to_value is None:
        cover.close()
        raise RuntimeError("No se localizo el valor To")

    exact_redaction(page, range_rect)
    exact_redaction(page, from_value)
    exact_redaction(page, to_value)
    page.apply_redactions()

    insert_original_style(page, range_rect, "Specific", style)
    insert_original_style(page, item_position, item, style)
    insert_original_style(page, revision_position, "A", style)
    return cover


def make_clean_source_page(source_doc, page_number, keep_y0, keep_y1):
    """Return a one-page doc with everything outside keep range removed."""
    temp = fitz.open()
    temp.insert_pdf(source_doc, from_page=page_number, to_page=page_number)
    page = temp[0]
    keep_y0 = max(page.rect.y0, float(keep_y0))
    keep_y1 = min(page.rect.y1, float(keep_y1))

    if keep_y0 > page.rect.y0:
        page.add_redact_annot(
            fitz.Rect(page.rect.x0, page.rect.y0, page.rect.x1, keep_y0),
            fill=(1, 1, 1),
        )
    if keep_y1 < page.rect.y1:
        page.add_redact_annot(
            fitz.Rect(page.rect.x0, keep_y1, page.rect.x1, page.rect.y1),
            fill=(1, 1, 1),
        )
    page.apply_redactions()
    return temp


def line_intervals(page, y0, y1):
    """Return merged occupied vertical intervals from words inside a range."""
    intervals = []
    for rect in visible_rects(page):
        if rect.y1 > y0 and rect.y0 < y1:
            intervals.append((max(y0, rect.y0), min(y1, rect.y1)))
    if not intervals:
        return []
    intervals.sort()
    merged = [list(intervals[0])]
    for start, end in intervals[1:]:
        if start <= merged[-1][1] + 1.5:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [(start, end) for start, end in merged]


def choose_safe_end(page, start_y, desired_end, absolute_end):
    """Choose a cut in whitespace at or before desired_end."""
    desired_end = min(desired_end, absolute_end)
    if desired_end >= absolute_end - 1:
        return absolute_end

    intervals = line_intervals(page, start_y, absolute_end)
    candidates = []
    previous_end = start_y
    for interval_start, interval_end in intervals:
        if interval_start > previous_end + 3:
            candidates.append((previous_end + interval_start) / 2)
        previous_end = max(previous_end, interval_end)
    if absolute_end > previous_end + 3:
        candidates.append((previous_end + absolute_end) / 2)

    valid = [value for value in candidates if start_y + MIN_SEGMENT_HEIGHT <= value <= desired_end]
    return max(valid) if valid else desired_end


def append_item_content(item_doc, source_doc, page_number, start_y, end_y):
    """Store a clean, searchable source segment with no next-item text."""
    clean = make_clean_source_page(source_doc, page_number, start_y, end_y)
    try:
        item_doc.insert_pdf(clean)
    finally:
        clean.close()


def content_bounds(page):
    """Vertical span of the content that is actually visible on the page."""
    rects = visible_rects(page)
    if not rects:
        return None
    return min(rect.y0 for rect in rects), max(rect.y1 for rect in rects)


def content_blocks(page):
    """Split a page's visible content into blocks, breaking at large gaps.

    Oracle leaves wide empty bands inside a routing page. Copying from the
    first to the last line drags those bands along, which is what produced the
    big blank areas in the output. Returning the blocks separately lets the
    composer place them one after another with a normal gap.
    """
    rects = visible_rects(page)
    if not rects:
        return []
    spans = sorted((rect.y0, rect.y1) for rect in rects)
    merged = [list(spans[0])]
    for start, end in spans[1:]:
        if start <= merged[-1][1] + 1.5:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])

    if MAX_BLANK_GAP <= 0:
        return [(merged[0][0], merged[-1][1])]

    blocks = [list(merged[0])]
    for start, end in merged[1:]:
        if start - blocks[-1][1] > MAX_BLANK_GAP:
            blocks.append([start, end])
        else:
            blocks[-1][1] = end
    return [(start, end) for start, end in blocks]


def compose_document(cover_doc, routing_doc):
    """
    Put routing directly below the cover and paginate at whitespace gaps.
    Orientation and page size are copied from the original cover.
    """
    final = fitz.open()
    final.insert_pdf(cover_doc)
    target = final[0]
    width, height = target.rect.width, target.rect.height

    cover_bounds = content_bounds(target)
    cursor_y = (cover_bounds[1] if cover_bounds else TOP_MARGIN) + COVER_TO_ROUTING_GAP

    for source_number in range(routing_doc.page_count):
        source_page = routing_doc[source_number]
        bloques = content_blocks(source_page)
        if not bloques:
            continue

        for indice, (source_y, source_end) in enumerate(bloques):
            if indice > 0:
                cursor_y += SOURCE_PAGE_GAP      # hueco acotado entre bloques

            while source_y < source_end - 0.5:
                available = height - BOTTOM_MARGIN - cursor_y
                if available < MIN_SEGMENT_HEIGHT:
                    target = final.new_page(width=width, height=height)
                    cursor_y = TOP_MARGIN
                    available = height - BOTTOM_MARGIN - cursor_y

                safe_end = choose_safe_end(source_page, source_y,
                                           source_y + available, source_end)
                if safe_end <= source_y + 1:
                    target = final.new_page(width=width, height=height)
                    cursor_y = TOP_MARGIN
                    continue

                clean_segment = make_clean_source_page(
                    routing_doc, source_number, source_y, safe_end)
                try:
                    clip = fitz.Rect(0, source_y, source_page.rect.width, safe_end)
                    destination = fitz.Rect(0, cursor_y, clip.width,
                                            cursor_y + clip.height)
                    target.show_pdf_page(destination, clean_segment, 0,
                                         clip=clip, overlay=True)
                finally:
                    clean_segment.close()

                cursor_y += safe_end - source_y
                source_y = safe_end

                if source_y < source_end - 0.5:
                    target = final.new_page(width=width, height=height)
                    cursor_y = TOP_MARGIN

        # Antes se abría SIEMPRE una página nueva al terminar cada página del
        # reporte original. Eso dejaba la mayor parte de la hoja en blanco
        # cuando la página fuente traía poco contenido. Ahora se continúa en la
        # misma hoja y solo se separa con un espacio.
        if source_number < routing_doc.page_count - 1:
            cursor_y += SOURCE_PAGE_GAP

    return final


def add_end_of_report(final_doc, source_doc):
    target = final_doc[-1]
    if "***** End of Report *****" in target.get_text("text"):
        return

    source_number = None
    footer_rect = None
    for number in range(source_doc.page_count - 1, -1, -1):
        hits = source_doc[number].search_for("***** End of Report *****")
        if hits:
            source_number, footer_rect = number, hits[-1]
            break

    if footer_rect is None:
        text = "***** End of Report *****"
        size = 8.2
        text_width = fitz.get_text_length(text, fontname="cour", fontsize=size)
        x = ((target.rect.width - text_width) / 2) + FOOTER_SHIFT_RIGHT
        target.insert_text((x, target.rect.height - 12), text, fontname="cour", fontsize=size)
        return

    source_page = source_doc[source_number]
    clip = fitz.Rect(
        max(0, footer_rect.x0 - 5), max(0, footer_rect.y0 - 3),
        min(source_page.rect.width, footer_rect.x1 + 5), min(source_page.rect.height, footer_rect.y1 + 3),
    )
    x = ((target.rect.width - clip.width) / 2) + FOOTER_SHIFT_RIGHT
    x = min(max(0, x), target.rect.width - clip.width)
    y = target.rect.height - clip.height - 8
    target.show_pdf_page(
        fitz.Rect(x, y, x + clip.width, y + clip.height),
        source_doc,
        source_number,
        clip=clip,
        overlay=True,
    )


def output_path(folder, item, source_stem):
    primary = folder / f"{safe_name(item)}.pdf"
    if not primary.exists():
        return primary
    return folder / f"{safe_name(item)}__{safe_name(source_stem)}.pdf"


def process_pdf(pdf_path, output_folder):
    source = fitz.open(pdf_path)
    item_docs = {}
    item_order = []
    start_pages = {}
    current_item = None

    try:
        for page_number in range(source.page_count):
            page = source[page_number]
            markers = item_markers(page)
            top, bottom = page.rect.y0, page.rect.y1

            if not markers:
                if current_item is not None:
                    append_item_content(item_docs[current_item], source, page_number, top, bottom)
                continue

            if current_item is not None and markers[0][1] > top + 3:
                append_item_content(item_docs[current_item], source, page_number, top, markers[0][1])

            for index, (item, start_y) in enumerate(markers):
                end_y = markers[index + 1][1] if index + 1 < len(markers) else bottom
                if item not in item_docs:
                    item_docs[item] = fitz.open()
                    item_order.append(item)
                    start_pages[item] = page_number + 1
                append_item_content(item_docs[item], source, page_number, start_y, end_y)
                current_item = item

        rows = []
        for item in item_order:
            routing = item_docs[item]
            cover = build_cover(source, item)
            final = None
            try:
                final = compose_document(cover, routing)
                add_end_of_report(final, source)

                full_text = "\n".join(final[number].get_text("text") for number in range(final.page_count))
                detected = {value.upper().rstrip(".,;:") for value in ITEM_RE.findall(full_text)}
                foreign = sorted(value for value in detected if value != item)
                if foreign:
                    raise RuntimeError(f"Items ajenos en {item}: {', '.join(foreign)}")

                destination = output_path(output_folder, item, pdf_path.stem)
                final.save(destination, garbage=4, deflate=True)
                rows.append({
                    "Part Number": item,
                    "Source PDF": pdf_path.name,
                    "Output PDF": destination.name,
                    "Final PDF Pages": final.page_count,
                    "Starts on Source Page": start_pages[item],
                    "Status": "Created",
                })
            finally:
                if final is not None:
                    final.close()
                cover.close()
                routing.close()
        return rows, None

    except Exception:
        for document in item_docs.values():
            try:
                document.close()
            except Exception:
                pass
        return [], traceback.format_exc()
    finally:
        source.close()


def create_summary(rows, path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Summary"
    headers = list(rows[0].keys())
    sheet.append(headers)

    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    for row in rows:
        sheet.append([row[header] for header in headers])

    widths = [22, 35, 40, 18, 24, 14]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def main():
    base = Path(__file__).resolve().parent
    input_folder = Path(sys.argv[1]) if len(sys.argv) > 1 else base / INPUT_FOLDER_NAME
    output_folder = Path(sys.argv[2]) if len(sys.argv) > 2 else base / OUTPUT_FOLDER_NAME
    input_folder.mkdir(parents=True, exist_ok=True)
    output_folder.mkdir(parents=True, exist_ok=True)
    pdfs = sorted(input_folder.glob("*.pdf"))

    if not pdfs:
        print(f"No se encontraron PDFs en: {input_folder}")
        return 1

    all_rows = []
    log = []
    errors = 0
    print(f"PDFs fuente: {len(pdfs)}")

    for index, pdf in enumerate(pdfs, start=1):
        print(f"[{index}/{len(pdfs)}] {pdf.name}")
        rows, error = process_pdf(pdf, output_folder)
        if error:
            errors += 1
            log.append(f"ERROR: {pdf.name}\n{error}")
            print("    ERROR")
        else:
            all_rows.extend(rows)
            log.append(f"OK: {pdf.name} -> {len(rows)} PDFs")
            print(f"    Creados: {len(rows)}")

    if all_rows:
        create_summary(all_rows, output_folder / SUMMARY_FILE_NAME)

    log.extend(["", f"Creados: {len(all_rows)}", f"Errores: {errors}"])
    (output_folder / LOG_FILE_NAME).write_text("\n".join(log), encoding="utf-8")
    print(f"\nCompletado. Creados: {len(all_rows)} | Errores: {errors}")
    return 0 if all_rows else 1


if __name__ == "__main__":
    raise SystemExit(main())
