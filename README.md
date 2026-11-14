# PDF Indexer

Automatically generates a clickable index for any PDF and prepends it to the document.

## What it does

- Scans every page of a PDF and extracts the topic/title from the top of the page
- Builds a two-column index with dotted leaders and page numbers
- Embeds clickable links in the index so each entry jumps directly to that page
- Adds bookmarks to the PDF outline
- Exports the index as Excel (`.xlsx`), Markdown (`.md`), HTML (`.html`), and CSV for review

## Usage

### Web app (recommended)

```bash
streamlit run app.py
```

Upload a PDF through the browser, then download the indexed PDF and all export formats directly.

### Command line

1. Drop your PDF(s) into the `input/` folder
2. Run:

```bash
python pdf_indexer.py
```

Outputs are saved to `outputs/<pdf-name>/`.

## Output files

| File | Description |
|---|---|
| `<name>_INDEXED.pdf` | Original PDF with the index prepended |
| `index.xlsx` | Index spreadsheet |
| `index.md` | Index in Markdown |
| `index.html` | Index as a web page |
| `review_pages.csv` | Pages where the title couldn't be detected confidently |

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

The main layout and detection settings are at the top of `pdf_indexer.py`:

| Setting | Default | Description |
|---|---|---|
| `TITLE_TOP_LIMIT` | `120` | Max Y position (px) to look for a title |
| `FALLBACK_TOP_LIMIT` | `180` | Fallback Y limit if nothing found in primary zone |
| `MIN_TITLE_LENGTH` | `4` | Minimum characters for a valid title |
| `MAX_TITLE_LENGTH` | `250` | Maximum characters for a valid title |
| `ROW_HEIGHT` | `20` | Row height in the index pages (px) |
