import pymupdf
import re
import csv
import html

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DictionaryObject,
    NameObject,
    ArrayObject,
    FloatObject,
)


# ============================================================
# FOLDERS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "outputs"

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# ============================================================
# TOPIC DETECTION SETTINGS
# ============================================================

# Your topic is at the top of the page.
TITLE_TOP_LIMIT = 120
FALLBACK_TOP_LIMIT = 180

MIN_TITLE_LENGTH = 4
MAX_TITLE_LENGTH = 250


# ============================================================
# INDEX PAGE SETTINGS
# ============================================================

PAGE_WIDTH = 595
PAGE_HEIGHT = 842

LEFT_MARGIN = 45
RIGHT_MARGIN = 45

TOP_MARGIN = 45
BOTTOM_MARGIN = 45

ROW_HEIGHT = 20

TITLE_FONT_SIZE = 18
SUBTITLE_FONT_SIZE = 9
TOPIC_FONT_SIZE = 8.5
PAGE_NUMBER_FONT_SIZE = 8.5

COLUMN_GAP = 25


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text(text):

    if not text:
        return ""

    replacements = {
        "\ufb00": "ff",
        "\ufb01": "fi",
        "\ufb02": "fl",
        "\ufb03": "ffi",
        "\ufb04": "ffl",
        "\u00ad": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# CHECK IF TEXT LOOKS LIKE A TOPIC
# ============================================================

def looks_like_title(text):

    text = clean_text(text)

    if len(text) < MIN_TITLE_LENGTH:
        return False

    if len(text) > MAX_TITLE_LENGTH:
        return False

    lower = text.lower()

    # Ignore URLs
    if lower.startswith("www."):
        return False

    if lower.startswith("http://"):
        return False

    if lower.startswith("https://"):
        return False

    # Ignore numbers / symbols only
    if re.fullmatch(
        r"[\d\s./_():,\-–—]+",
        text
    ):
        return False

    return True


# ============================================================
# GET TEXT BLOCKS FROM PAGE
# ============================================================

def get_text_blocks(page):

    blocks = []

    page_dict = page.get_text("dict")

    for block in page_dict.get("blocks", []):

        # type 0 = text
        if block.get("type") != 0:
            continue

        lines = block.get("lines", [])

        text_parts = []
        font_sizes = []

        for line in lines:

            for span in line.get("spans", []):

                text = span.get("text", "")

                if text:
                    text_parts.append(text)

                size = span.get("size")

                if size:
                    font_sizes.append(
                        float(size)
                    )

        text = clean_text(
            " ".join(text_parts)
        )

        if not text:
            continue

        x0, y0, x1, y1 = block["bbox"]

        blocks.append(
            {
                "text": text,
                "x0": x0,
                "y0": y0,
                "x1": x1,
                "y1": y1,
                "font_size": (
                    max(font_sizes)
                    if font_sizes
                    else 0
                ),
            }
        )

    return blocks


# ============================================================
# EXTRACT TOPIC FROM PAGE
# ============================================================

def extract_title(page):

    blocks = get_text_blocks(page)

    # --------------------------------------------------------
    # PRIMARY SEARCH
    # --------------------------------------------------------

    candidates = []

    for block in blocks:

        if block["y0"] > TITLE_TOP_LIMIT:
            continue

        if looks_like_title(
            block["text"]
        ):

            candidates.append(block)

    if candidates:

        # Prefer text that is:
        # 1. Near the top
        # 2. Larger font
        #
        # This is useful for your layout where
        # the topic is at the top-right.

        candidates.sort(
            key=lambda block: (
                block["y0"],
                -block["font_size"],
                block["x0"]
            )
        )

        return (
            candidates[0]["text"],
            "FOUND"
        )

    # --------------------------------------------------------
    # FALLBACK
    # --------------------------------------------------------

    candidates = []

    for block in blocks:

        if block["y0"] > FALLBACK_TOP_LIMIT:
            continue

        if looks_like_title(
            block["text"]
        ):

            candidates.append(block)

    if candidates:

        candidates.sort(
            key=lambda block: (
                block["y0"],
                -block["font_size"],
                block["x0"]
            )
        )

        return (
            candidates[0]["text"],
            "FALLBACK"
        )

    return (
        "TITLE NOT FOUND — REVIEW PAGE",
        "REVIEW"
    )


# ============================================================
# EXTRACT TOPICS FROM ENTIRE PDF
# ============================================================

def extract_topics(pdf_path):

    print()
    print("=" * 70)
    print("EXTRACTING TOPICS")
    print("=" * 70)

    document = pymupdf.open(
        pdf_path
    )

    total_pages = len(document)

    print(
        f"PDF: {pdf_path.name}"
    )

    print(
        f"Total pages: {total_pages}"
    )

    print()

    results = []

    for page_number in range(
        1,
        total_pages + 1
    ):

        page = document[
            page_number - 1
        ]

        title, status = extract_title(
            page
        )

        results.append(
            {
                "page": page_number,
                "title": title,
                "status": status
            }
        )

        if (
            page_number == 1
            or page_number % 50 == 0
            or page_number == total_pages
        ):

            print(
                f"Processed "
                f"{page_number}/{total_pages}"
            )

    document.close()

    found = sum(
        item["status"] == "FOUND"
        for item in results
    )

    fallback = sum(
        item["status"] == "FALLBACK"
        for item in results
    )

    review = sum(
        item["status"] == "REVIEW"
        for item in results
    )

    print()
    print(
        f"Normal titles : {found}"
    )
    print(
        f"Fallback      : {fallback}"
    )
    print(
        f"Review needed : {review}"
    )

    return results


# ============================================================
# CREATE EXCEL INDEX
# ============================================================

def create_excel(
    results,
    output_path
):

    print(
        "Creating index.xlsx..."
    )

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Index"

    headers = [
        "Original Page",
        "Topic",
        "Status"
    ]

    sheet.append(headers)

    for cell in sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="D9EAF7"
        )

    for item in results:

        sheet.append(
            [
                item["page"],
                item["title"],
                item["status"]
            ]
        )

    sheet.column_dimensions[
        "A"
    ].width = 18

    sheet.column_dimensions[
        "B"
    ].width = 80

    sheet.column_dimensions[
        "C"
    ].width = 18

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = (
        sheet.dimensions
    )

    workbook.save(
        output_path
    )


# ============================================================
# CREATE MARKDOWN
# ============================================================

def create_markdown(
    results,
    output_path,
    pdf_name
):

    print(
        "Creating index.md..."
    )

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            f"# Index — {pdf_name}\n\n"
        )

        file.write(
            f"Total pages: {len(results)}\n\n"
        )

        file.write(
            "| Page | Topic | Status |\n"
        )

        file.write(
            "|---:|---|---|\n"
        )

        for item in results:

            topic = (
                item["title"]
                .replace("|", "\\|")
            )

            file.write(
                f"| {item['page']} | "
                f"{topic} | "
                f"{item['status']} |\n"
            )


# ============================================================
# CREATE HTML
# ============================================================

def create_html(
    results,
    output_path,
    pdf_name
):

    print(
        "Creating index.html..."
    )

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            """
<!DOCTYPE html>
<html>
<head>

<meta charset="UTF-8">

<title>PDF Index</title>

<style>

body {
    font-family: Arial, sans-serif;
    margin: 40px;
}

table {
    border-collapse: collapse;
    width: 100%;
}

th, td {
    border-bottom: 1px solid #ddd;
    padding: 10px;
    text-align: left;
}

th {
    background: #f2f2f2;
}

</style>

</head>

<body>
"""
        )

        file.write(
            "<h1>Index — "
            + html.escape(pdf_name)
            + "</h1>"
        )

        file.write(
            f"<p>{len(results)} pages</p>"
        )

        file.write(
            """
<table>

<tr>
<th>Page</th>
<th>Topic</th>
<th>Status</th>
</tr>
"""
        )

        for item in results:

            file.write(
                f"""
<tr>

<td>{item["page"]}</td>

<td>{html.escape(item["title"])}</td>

<td>{item["status"]}</td>

</tr>
"""
            )

        file.write(
            """
</table>

</body>

</html>
"""
        )


# ============================================================
# CALCULATE INDEX PAGE CHUNKS
# ============================================================

def calculate_index_chunks(results):

    first_page_rows = int(
        (
            PAGE_HEIGHT
            - TOP_MARGIN
            - BOTTOM_MARGIN
            - 80
        )
        / ROW_HEIGHT
    )

    normal_page_rows = int(
        (
            PAGE_HEIGHT
            - TOP_MARGIN
            - BOTTOM_MARGIN
            - 45
        )
        / ROW_HEIGHT
    )

    chunks = []

    current = []

    capacity = first_page_rows

    for item in results:

        current.append(item)

        if len(current) >= capacity:

            chunks.append(current)

            current = []

            capacity = normal_page_rows

    if current:

        chunks.append(current)

    if not chunks:

        chunks.append([])

    return chunks


# ============================================================
# CREATE INDEX PDF
# ============================================================

def create_index_pdf(
    results,
    pdf_name,
    index_pdf_path
):

    print()
    print("=" * 70)
    print("CREATING INDEX PDF")
    print("=" * 70)

    document = pymupdf.open()

    chunks = calculate_index_chunks(
        results
    )

    # Store the clickable row information.
    #
    # Each entry:
    #
    # {
    #     "index_page": 0,
    #     "original_page": 1,
    #     "rect": (...)
    # }

    link_data = []

    for index_number, chunk in enumerate(
        chunks
    ):

        page = document.new_page(
            width=PAGE_WIDTH,
            height=PAGE_HEIGHT
        )

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        if index_number == 0:

            page.insert_text(
                (
                    LEFT_MARGIN,
                    TOP_MARGIN + 18
                ),
                "INDEX",
                fontsize=TITLE_FONT_SIZE,
                fontname="helv"
            )

            page.insert_text(
                (
                    LEFT_MARGIN,
                    TOP_MARGIN + 38
                ),
                pdf_name,
                fontsize=SUBTITLE_FONT_SIZE,
                fontname="helv"
            )

            page.insert_text(
                (
                    LEFT_MARGIN,
                    TOP_MARGIN + 52
                ),
                f"{len(results)} original pages",
                fontsize=7.5,
                fontname="helv"
            )

            header_y = TOP_MARGIN + 75

        else:

            page.insert_text(
                (
                    LEFT_MARGIN,
                    TOP_MARGIN + 18
                ),
                "INDEX",
                fontsize=12,
                fontname="helv"
            )

            page.insert_text(
                (
                    PAGE_WIDTH
                    - RIGHT_MARGIN
                    - 20,
                    TOP_MARGIN + 18
                ),
                str(index_number + 1),
                fontsize=8,
                fontname="helv"
            )

            header_y = TOP_MARGIN + 45

        # ----------------------------------------------------
        # TWO COLUMNS
        # ----------------------------------------------------

        split = (
            len(chunk) + 1
        ) // 2

        columns = [
            chunk[:split],
            chunk[split:]
        ]

        column_width = (
            PAGE_WIDTH
            - LEFT_MARGIN
            - RIGHT_MARGIN
            - COLUMN_GAP
        ) / 2

        for column_number, records in enumerate(
            columns
        ):

            column_left = (
                LEFT_MARGIN
                + column_number
                * (
                    column_width
                    + COLUMN_GAP
                )
            )

            column_right = (
                column_left
                + column_width
            )

            y = header_y

            for item in records:

                topic = clean_text(
                    item["title"]
                )

                if len(topic) > 65:

                    topic = (
                        topic[:64]
                        + "…"
                    )

                page_number = str(
                    item["page"]
                )

                # ------------------------------------------------
                # PAGE NUMBER WIDTH
                # ------------------------------------------------

                font = pymupdf.Font(
                    "helv"
                )

                page_number_width = (
                    font.text_length(
                        page_number,
                        fontsize=PAGE_NUMBER_FONT_SIZE
                    )
                )

                # ------------------------------------------------
                # TOPIC
                # ------------------------------------------------

                page.insert_text(
                    (
                        column_left,
                        y
                    ),
                    topic,
                    fontsize=TOPIC_FONT_SIZE,
                    fontname="helv"
                )

                # ------------------------------------------------
                # PAGE NUMBER
                # ------------------------------------------------

                page_number_x = (
                    column_right
                    - page_number_width
                )

                page.insert_text(
                    (
                        page_number_x,
                        y
                    ),
                    page_number,
                    fontsize=PAGE_NUMBER_FONT_SIZE,
                    fontname="helv"
                )

                # ------------------------------------------------
                # DOTTED LEADER
                # ------------------------------------------------

                topic_width = (
                    font.text_length(
                        topic,
                        fontsize=TOPIC_FONT_SIZE
                    )
                )

                dots_start = (
                    column_left
                    + topic_width
                    + 5
                )

                dots_end = (
                    page_number_x
                    - 5
                )

                if dots_end > dots_start:

                    dot_count = int(
                        (
                            dots_end
                            - dots_start
                        )
                        / 4
                    )

                    if dot_count > 0:

                        dots = "." * dot_count

                        page.insert_text(
                            (
                                dots_start,
                                y
                            ),
                            dots,
                            fontsize=TOPIC_FONT_SIZE,
                            fontname="helv"
                        )

                # ------------------------------------------------
                # CLICKABLE RECTANGLE
                # ------------------------------------------------
                #
                # PyMuPDF uses top-left coordinates.
                #
                # This rectangle covers the whole row.
                #

                row_rect = pymupdf.Rect(
                    column_left - 3,
                    y - ROW_HEIGHT + 4,
                    column_right + 3,
                    y + 5
                )

                link_data.append(
                    {
                        "index_page": index_number,
                        "original_page": item["page"],
                        "rect": (
                            float(row_rect.x0),
                            float(row_rect.y0),
                            float(row_rect.x1),
                            float(row_rect.y1)
                        )
                    }
                )

                y += ROW_HEIGHT

    # --------------------------------------------------------
    # SAVE INDEX PDF
    # --------------------------------------------------------

    document.save(
        index_pdf_path
    )

    document.close()

    print(
        f"Index pages created: "
        f"{len(chunks)}"
    )

    return (
        len(chunks),
        link_data
    )


# ============================================================
# ADD CLICKABLE LINKS WITH PYPDF
# ============================================================

def add_clickable_links(
    writer,
    link_data,
    index_page_count
):

    print(
        "Adding clickable index links..."
    )

    for link in link_data:

        index_page = link[
            "index_page"
        ]

        original_page = link[
            "original_page"
        ]

        # ----------------------------------------------------
        # Original page 1 becomes:
        #
        # index_page_count + 0
        #
        # Original page 2 becomes:
        #
        # index_page_count + 1
        # ----------------------------------------------------

        target_page = (
            index_page_count
            + original_page
            - 1
        )

        if (
            target_page < 0
            or target_page >= len(writer.pages)
        ):

            continue

        x0, y0, x1, y1 = link[
            "rect"
        ]

        page = writer.pages[
            index_page
        ]

        page_height = float(
            page.mediabox.height
        )

        # ----------------------------------------------------
        # PyMuPDF:
        #
        # origin = TOP LEFT
        #
        # pypdf:
        #
        # origin = BOTTOM LEFT
        #
        # Therefore convert Y.
        # ----------------------------------------------------

        pypdf_y0 = (
            page_height - y1
        )

        pypdf_y1 = (
            page_height - y0
        )

        annotation = DictionaryObject()

        annotation.update(
            {
                NameObject("/Type"):
                    NameObject("/Annot"),

                NameObject("/Subtype"):
                    NameObject("/Link"),

                NameObject("/Rect"):
                    ArrayObject(
                        [
                            FloatObject(x0),
                            FloatObject(pypdf_y0),
                            FloatObject(x1),
                            FloatObject(pypdf_y1),
                        ]
                    ),

                NameObject("/Border"):
                    ArrayObject(
                        [
                            FloatObject(0),
                            FloatObject(0),
                            FloatObject(0),
                        ]
                    ),

                NameObject("/Dest"):
                    ArrayObject(
                        [
                            writer.pages[
                                target_page
                            ].indirect_reference,

                            NameObject("/Fit")
                        ]
                    )
            }
        )

        page[NameObject(
            "/Annots"
        )] = page.get(
            "/Annots",
            ArrayObject()
        )

        page["/Annots"].append(
            annotation
        )


# ============================================================
# COMBINE INDEX + ORIGINAL
# ============================================================

def combine_with_pypdf(
    index_pdf_path,
    original_pdf_path,
    final_pdf_path,
    link_data
):

    print()
    print("=" * 70)
    print("COMBINING PDF USING PYPDF")
    print("=" * 70)

    print(
        "Opening index PDF..."
    )

    index_reader = PdfReader(
        str(index_pdf_path),
        strict=False
    )

    print(
        "Opening original PDF..."
    )

    original_reader = PdfReader(
        str(original_pdf_path),
        strict=False
    )

    index_page_count = len(
        index_reader.pages
    )

    original_page_count = len(
        original_reader.pages
    )

    print(
        f"Index pages: "
        f"{index_page_count}"
    )

    print(
        f"Original pages: "
        f"{original_page_count}"
    )

    writer = PdfWriter()

    # --------------------------------------------------------
    # Add index pages
    # --------------------------------------------------------

    print(
        "Adding index pages..."
    )

    for page in index_reader.pages:

        writer.add_page(
            page
        )

    # --------------------------------------------------------
    # Add original pages
    # --------------------------------------------------------

    print(
        "Adding original pages..."
    )

    for number, page in enumerate(
        original_reader.pages,
        start=1
    ):

        writer.add_page(
            page
        )

        if (
            number == 1
            or number % 100 == 0
            or number == original_page_count
        ):

            print(
                f"Added original page "
                f"{number}/"
                f"{original_page_count}"
            )

    # --------------------------------------------------------
    # Add clickable links
    # --------------------------------------------------------

    add_clickable_links(
        writer,
        link_data,
        index_page_count
    )

    # --------------------------------------------------------
    # Add bookmarks
    # --------------------------------------------------------

    print(
        "Adding bookmarks..."
    )

    writer.add_outline_item(
        "INDEX",
        0
    )

    for item in link_data:

        original_page = item[
            "original_page"
        ]

        target_page = (
            index_page_count
            + original_page
            - 1
        )

        # Only add one bookmark per original page
        #
        # link_data has one entry per index row.

        if target_page < len(
            writer.pages
        ):

            title = (
                f"Page {original_page} — "
                f"{get_topic_for_page(
                    original_page
                )}"
            )

            if len(title) > 180:

                title = (
                    title[:179]
                    + "…"
                )

            writer.add_outline_item(
                title,
                target_page
            )

    # --------------------------------------------------------
    # Write final PDF
    # --------------------------------------------------------

    print()
    print(
        "Writing final PDF..."
    )

    with open(
        final_pdf_path,
        "wb"
    ) as output_file:

        writer.write(
            output_file
        )

    print()
    print(
        "Final PDF successfully created."
    )

    print(
        final_pdf_path
    )

    return index_page_count


# ============================================================
# GLOBAL TOPIC LOOKUP FOR BOOKMARKS
# ============================================================

TOPIC_LOOKUP = {}


def get_topic_for_page(page_number):

    return TOPIC_LOOKUP.get(
        page_number,
        "Topic"
    )


# ============================================================
# REVIEW CSV
# ============================================================

def create_review_csv(
    results,
    output_path
):

    print(
        "Creating review_pages.csv..."
    )

    with open(
        output_path,
        "w",
        newline="",
        encoding="utf-8"
    ) as file:

        writer = csv.writer(
            file
        )

        writer.writerow(
            [
                "Original Page",
                "Topic",
                "Status"
            ]
        )

        for item in results:

            if item["status"] != "FOUND":

                writer.writerow(
                    [
                        item["page"],
                        item["title"],
                        item["status"]
                    ]
                )


# ============================================================
# PROCESS ONE PDF
# ============================================================

def process_pdf(
    pdf_path
):

    global TOPIC_LOOKUP

    print()
    print("=" * 70)
    print(
        f"STARTING: {pdf_path.name}"
    )
    print("=" * 70)

    # --------------------------------------------------------
    # Output folder
    # --------------------------------------------------------

    output_folder = (
        OUTPUT_DIR
        / pdf_path.stem
    )

    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Extract topics
    # --------------------------------------------------------

    results = extract_topics(
        pdf_path
    )

    # Build lookup for bookmarks
    TOPIC_LOOKUP = {
        item["page"]: item["title"]
        for item in results
    }

    # --------------------------------------------------------
    # Create Excel
    # --------------------------------------------------------

    create_excel(
        results,
        output_folder
        / "index.xlsx"
    )

    # --------------------------------------------------------
    # Create Markdown
    # --------------------------------------------------------

    create_markdown(
        results,
        output_folder
        / "index.md",
        pdf_path.name
    )

    # --------------------------------------------------------
    # Create HTML
    # --------------------------------------------------------

    create_html(
        results,
        output_folder
        / "index.html",
        pdf_path.name
    )

    # --------------------------------------------------------
    # Create temporary index PDF
    # --------------------------------------------------------

    temporary_index_pdf = (
        output_folder
        / "_index_only.pdf"
    )

    (
        index_page_count,
        link_data
    ) = create_index_pdf(
        results,
        pdf_path.name,
        temporary_index_pdf
    )

    # --------------------------------------------------------
    # Final PDF
    # --------------------------------------------------------

    final_pdf = (
        output_folder
        / f"{pdf_path.stem}_INDEXED.pdf"
    )

    # --------------------------------------------------------
    # Combine
    # --------------------------------------------------------

    combine_with_pypdf(
        temporary_index_pdf,
        pdf_path,
        final_pdf,
        link_data
    )

    # --------------------------------------------------------
    # Review CSV
    # --------------------------------------------------------

    create_review_csv(
        results,
        output_folder
        / "review_pages.csv"
    )

    # --------------------------------------------------------
    # Delete temporary PDF
    # --------------------------------------------------------

    if temporary_index_pdf.exists():

        temporary_index_pdf.unlink()

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("SUCCESS")
    print("=" * 70)

    print(
        f"Original pages : {len(results)}"
    )

    print(
        f"Index pages    : {index_page_count}"
    )

    print(
        f"Final pages    : "
        f"{len(results) + index_page_count}"
    )

    print()
    print(
        "FINAL PDF:"
    )

    print(
        final_pdf
    )

    print()


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print("PDF TOPIC INDEXER")
    print("=" * 70)

    # --------------------------------------------------------
    # Find PDFs automatically
    # --------------------------------------------------------

    pdf_files = sorted(
        INPUT_DIR.glob("*.pdf")
    )

    if not pdf_files:

        print()
        print(
            "ERROR: No PDF found."
        )

        print()
        print(
            "Put your PDF inside:"
        )

        print(
            INPUT_DIR
        )

        return

    print()
    print(
        f"Found {len(pdf_files)} PDF(s):"
    )

    for pdf in pdf_files:

        print(
            f"  • {pdf.name}"
        )

    print()

    # --------------------------------------------------------
    # Process each PDF
    # --------------------------------------------------------

    for pdf_path in pdf_files:

        try:

            process_pdf(
                pdf_path
            )

        except Exception as error:

            print()
            print("=" * 70)
            print(
                f"ERROR processing "
                f"{pdf_path.name}"
            )
            print(
                f"{type(error).__name__}: "
                f"{error}"
            )
            print("=" * 70)

    print()
    print("=" * 70)
    print("PROCESSING FINISHED")
    print("=" * 70)
    print()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()