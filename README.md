2# File Format Converter

A simple GUI application for converting files between different formats. This application supports various image and data file format conversions.

## Supported Conversions

### Image Formats
- PNG ↔ JPG/JPEG
- PNG ↔ BMP
- PNG ↔ GIF
- JPG/JPEG ↔ BMP
- JPG/JPEG ↔ GIF
- BMP ↔ GIF

### Data Formats
- CSV ↔ JSON
- CSV ↔ XLSX
- JSON ↔ XLSX

## Installation

1. Clone this repository or download the files
2. Create a virtual environment (recommended):
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```
2. Click "Browse" to select a file to convert
3. Choose the target format from the dropdown menu
4. Click "Convert" to start the conversion
5. The converted file will be saved in the same directory as the source file

## Features

- Simple and intuitive user interface
- Progress bar for conversion status
- Error handling and user feedback
- Support for multiple file formats
- Automatic format detection 