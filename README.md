# File Converter

A powerful Dify plugin providing comprehensive local file conversion capabilities. Supports conversion between PDF, Image, Word, Excel, PowerPoint, and Text formats with high-quality output and flexible options.

## Version Information

- **Current Version**: v0.0.2
- **Release Date**: 2025-12-28
- **Compatibility**: Dify Plugin Framework
- **Python Version**: 3.12

### Version History
- **v0.0.2** (2025-12-28): Added CSV to Excel, Excel to CSV, and CSV to PDF conversion capabilities with smart layout optimization
- **v0.0.1** (2025-11-02): Initial release with comprehensive file conversion capabilities

## Quick Start

1. Download the file_converter plugin from the Dify marketplace
2. Install the plugin in your Dify environment
3. Start using the various conversion tools immediately

## Core Features
<img width="415" height="954" alt="image" src="https://github.com/user-attachments/assets/2a818093-bba6-4196-84e1-23186aaac2ca" />

### PDF Conversions
- **PDF to Image**: Convert PDF documents to various image formats (.jpg, .jpeg, .png, .bmp, .tiff)
- **PDF to Word**: Extract content from PDF and convert to Word documents
- **PDF to Text**: Extract text content from PDF documents

### Image Conversions
- **Image to PDF**: Convert one or more images to PDF format

### Microsoft Office Conversions
- **Word to PDF**: Convert Word documents to PDF format
- **Word to Text**: Extract text content from Word documents
- **Excel to PDF**: Convert Excel spreadsheets to PDF format
- **PowerPoint to PDF**: Convert PowerPoint presentations to PDF format

### Text Conversions
- **Text to PDF**: Convert plain text files to PDF format
- **Text to Word**: Convert plain text files to Word documents

### CSV Conversions
- **CSV to Excel**: Convert CSV files to Excel format with automatic column width adjustment
- **Excel to CSV**: Convert Excel files to CSV format, supporting all worksheets
- **CSV to PDF**: Convert CSV files to PDF format with smart layout optimization and automatic column width adjustment

## Technical Advantages

- **Local Processing**: All conversions are performed locally without external dependencies
- **High-Quality Output**: Maintains original quality and formatting as much as possible
- **Multiple Format Support**: Comprehensive support for common file formats
- **Error Handling**: Robust error handling with informative messages
- **Flexible Options**: Various configuration options for different use cases
- **Secure Processing**: Files are processed securely without data retention

## Requirements

- Python 3.12
- Dify Platform access
- Required Python packages (installed via requirements.txt)

## Installation & Configuration

1. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Install the plugin in your Dify environment following the standard plugin installation process

## Usage

The plugin provides various tools for different file conversion tasks:

### PDF Conversions

#### 1. PDF to Image (pdf_2_image)
Convert PDF documents to image format.
- **Parameters**:
  - `input_file`: The PDF document to convert (required)
  - `output_format`: Output image format (jpg, jpeg, png, bmp, or tiff, default: png)

#### 2. PDF to Word (pdf_2_word)
Convert PDF documents to Word format.
- **Parameters**:
  - `input_file`: The PDF document to convert (required)

#### 3. PDF to Text (pdf_2_text)
Extract text content from PDF documents.
- **Parameters**:
  - `input_file`: The PDF document to extract text from (required)

### Image Conversions

#### 4. Image to PDF (image_2_pdf)
Convert one or more images to a single PDF document.
- **Parameters**:
  - `input_files`: The image files to convert (required, supports multiple files)
- **Features**:
  - Supports multiple image formats (.jpg, .jpeg, .png, .bmp, .tiff)
  - Converts multiple images into a single PDF document
  - Maintains the order of images as they are uploaded
  - Each image becomes a separate page in the PDF

### Microsoft Office Conversions

#### 5. Word to PDF (word_2_pdf)
Convert Word documents to PDF format.
- **Parameters**:
  - `input_file`: The Word document to convert (required)

#### 6. Word to Text (word_2_text)
Extract text content from Word documents.
- **Parameters**:
  - `input_file`: The Word document to extract text from (required)

#### 7. Excel to PDF (excel_2_pdf)
Convert Excel spreadsheets to PDF format.
- **Parameters**:
  - `input_file`: The Excel file to convert (required)
- **Features**:
  - Supports both .xlsx and .xls formats
  - Smart layout optimization for different table sizes
  - Automatic column width adjustment based on content
  - Landscape orientation for wide tables
  - Table splitting for excessively wide tables
  - Font scaling to fit content within page boundaries

#### 8. PowerPoint to PDF (ppt_2_pdf)
Convert PowerPoint presentations to PDF format.
- **Parameters**:
  - `input_file`: The PowerPoint file to convert (required)

### Text Conversions

#### 9. Text to PDF (text_2_pdf)
Convert plain text files to PDF format.
- **Parameters**:
  - `input_file`: The text file to convert (required)

#### 10. Text to Word (text_2_word)
Convert plain text files to Word documents.
- **Parameters**:
  - `input_file`: The text file to convert (required)

### CSV Conversions

#### 11. CSV to Excel (csv_2_excel)
Convert CSV files to Excel format.
- **Parameters**:
  - `input_file`: The CSV file to convert (required)
- **Features**:
  - Supports multiple encodings (utf-8, gbk, gb2312, latin-1, iso-8859-1)
  - Automatic column width adjustment
  - Sanitized sheet names to meet Excel specifications

#### 12. Excel to CSV (excel_2_csv)
Convert Excel files to CSV format.
- **Parameters**:
  - `input_file`: The Excel file to convert (required)
- **Features**:
  - Supports all worksheets in the Excel file (.xlsx, .xls)
  - Each worksheet is converted to a separate CSV file
  - Maintains data integrity and formatting

#### 13. CSV to PDF (csv_2_pdf)
Convert CSV files to PDF format with smart layout optimization.
- **Parameters**:
  - `input_file`: The CSV file to convert (required)
- **Features**:
  - Smart layout optimization for different table sizes
  - Automatic column width adjustment based on content
  - Landscape orientation for wide tables
  - Table splitting for excessively wide tables
  - Font scaling to fit content within page boundaries
  - Multiple encoding support (utf-8, gbk, gb2312, latin-1, iso-8859-1)

## Notes

- All conversions are performed locally without uploading files to external services
- Some conversions may require additional libraries that are included in the requirements.txt
- Large files may take longer to process depending on their complexity and size
- The quality of output files depends on the quality and format of the input files

## Developer Information

- **Author**: `https://github.com/sawyer-shi`
- **Email**: sawyer36@foxmail.com
- **License**: Apache License 2.0
- **Source Code**: `https://github.com/sawyer-shi/dify-plugins-file_converter`
- **Support**: Through Dify platform and GitHub Issues

## License Notice

This project is licensed under the Apache License 2.0. See the [LICENSE](LICENSE) file for the full license text.

**Note**: This project was previously licensed under MIT License but has been updated to Apache License 2.0 starting from version 0.0.2.

---

**Ready to convert your files with ease?**



