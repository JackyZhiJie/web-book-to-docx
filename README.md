# Web Book to DOCX Converter

This is a Python script that automatically converts an entire online web book into a single, perfectly formatted Microsoft Word (`.docx`) document. 

Unlike standard web scrapers that lose formatting, CSS layouts, and images, this script leverages Microsoft Word's native HTML rendering engine via COM automation. This ensures that the generated document is visually identical to the webpage—preserving all images, tables, grid layouts, and styles.

## Features

- **Perfect Formatting:** Uses MS Word's engine to render the HTML, guaranteeing pixel-perfect layouts.
- **Smart Chapter Extraction:** Automatically parses the main book URL to find sub-links (chapters) that belong to the book.
- **Automated Merging:** Downloads each chapter and merges them sequentially into a single master document, inserting page breaks between chapters.
- **Crash-Proof & Fast:** Uses the native `InsertFile` method instead of background clipboard operations, preventing COM object crashes and drastically speeding up the conversion.

## Requirements

- **Operating System:** Windows
- **Software:** Microsoft Word must be installed on your system.
- **Python:** Python 3.6 or newer.

## Installation

1. Clone or download this repository.
2. Install the required Python package (`pywin32`):

```bash
pip install pywin32
```

## Usage

Run the script from your terminal/command prompt by providing the URL to the main page of the book:

```bash
python converter.py "https://www.example.com/book/" -o "My Book Title.docx"
```

### Arguments:
- `url`: (Required) The link to the main page or table of contents of the book.
- `-o`, `--output`: (Optional) The output filename. Defaults to `output.docx`.

## How it works

1. **Extraction:** The script sends a standard HTTP request to the provided URL and parses the HTML to find all internal `.html` links that belong to the same domain.
2. **Bootstrapping:** It launches Microsoft Word silently in the background.
3. **Merging:** It loops through each chapter URL and instructs Word to directly insert and render the HTML from the web into the master document.
<<<<<<< Updated upstream
4. **Saving:** Once all chapters are inserted, it saves the file as `.docx` and closes Microsoft Word.
=======
4. **Saving:** Once all chapters are inserted, it saves the file as `.docx` and closes Microsoft Word.
>>>>>>> Stashed changes
