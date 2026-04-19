# Batch Document to Markdown Tool

A cross-platform desktop application that batch converts DOCX and PPTX files to Markdown format with a user-friendly GUI.

## Features

- **Batch Conversion**: Convert multiple DOCX and PPTX files simultaneously
- **Recursive Processing**: Automatically scans all subdirectories within the source folder
- **Directory Preservation**: Maintains the original folder structure in the output directory
- **Image Extraction**: Extracts images from documents and saves them to an `img` subfolder
- **Rich Text Support**: Preserves bold, italic, underline formatting
- **Table Support**: Converts tables to Markdown table syntax
- **Progress Tracking**: Real-time conversion progress display
- **User-Friendly GUI**: Simple and intuitive interface

## Supported Formats

| Input Format | Description |
|-------------|-------------|
| `.docx` | Microsoft Word documents |
| `.pptx` | Microsoft PowerPoint presentations |

## Installation

### Prerequisites

- Python 3.8+
- pip (Python package manager)

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Run the Application

```bash
python main.py
```

## Usage

### Interface Overview

1. **Source Directory**: Select the folder containing DOCX and PPTX files to convert
2. **Target Directory**: Choose where to save the converted Markdown files
3. **Start Processing**: Click to begin batch conversion
4. **Progress Display**: View real-time conversion status and results

### Example

```
Source Directory:  C:\Users\Documents\MyProject
Target Directory:  C:\Users\Output\MyProject_MD
```

After conversion:
```
C:\Users\Output\MyProject_MD\
├── document1.md
├── document2.md
├── subfolder\
│   ├── presentation1.md
│   └── notes.md
└── img\
    ├── document1_image_0.png
    └── presentation1_image_0.png
```

## Conversion Details

### DOCX Conversion

- **Headings**: Converts Word heading styles (Heading 1-6) to Markdown headers (# to ######)
- **Paragraphs**: Preserves plain text paragraphs
- **Lists**: Converts bulleted and numbered lists to Markdown list syntax
- **Tables**: Converts Word tables to Markdown table syntax
- **Images**: Extracts inline images, preserves transparency as white background
- **JSON Blocks**: Handles JSON content within documents
- **Text Formatting**: Preserves bold, italic, and underline styling

### PPTX Conversion

- **Slides**: Processes PowerPoint slides as separate content blocks
- **Text Content**: Extracts title and body text
- **Lists**: Converts bullet points to Markdown lists
- **Tables**: Converts PowerPoint tables to Markdown tables
- **Images**: Extracts embedded images with proper file handling
- **WMF Images**: Converts Windows Metafile format to PNG on Windows

## Project Structure

```
docx_pptx_to_md_en/
├── main.py                 # Application entry point
├── gui.py                  # GUI interface implementation
├── docx_converter.py       # DOCX to Markdown converter
├── pptx_converter.py       # PPTX to Markdown converter
├── parser.py               # PPTX parsing logic
├── outputter.py            # Markdown output generation
├── custom_types.py         # Data type definitions
├── entry.py               # Conversion pipeline
├── multi_column.py         # Multi-column slide detection
├── utils.py                # Utility functions
├── log.py                  # Logging configuration
└── requirements.txt        # Python dependencies
```

## Dependencies

- `python-docx` - DOCX file parsing
- `python-pptx` - PPTX file parsing
- `Pillow` - Image processing
- `wand` - WMF image conversion (optional)
- `rapidfuzz` - Fuzzy string matching for title processing
- `tqdm` - Progress bar display

## Building Executable

To create a standalone executable (.exe) for Windows:

```bash
pip install pyinstaller
pyinstaller docx_pptx_to_md.spec
```

The executable will be generated in the `dist` folder.

## License

Apache License 2.0

## Author

Li Yong
