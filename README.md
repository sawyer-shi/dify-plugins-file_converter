# File Converter

A powerful Dify plugin providing comprehensive local file conversion capabilities. Supports conversion between PDF, Image, Word, Excel, PowerPoint, and Text formats with high-quality output and flexible options.

## Version Information

- **Current Version**: v0.0.1
- **Release Date**: 2025-11-02
- **Compatibility**: Dify Plugin Framework
- **Python Version**: 3.12

### Version History
- **v0.0.1** (2025-11-02): Initial release with comprehensive file conversion capabilities

## Quick Start

1. Download the file_converter plugin from the Dify marketplace
2. Install the plugin in your Dify environment
3. Start using the various conversion tools immediately

## Core Features

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
Convert images to PDF format.
- **Parameters**:
  - `input_file`: The image file to convert (required)

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

## Notes

- All conversions are performed locally without uploading files to external services
- Some conversions may require additional libraries that are included in the requirements.txt
- Large files may take longer to process depending on their complexity and size
- The quality of output files depends on the quality and format of the input files

## Developer Information

- **Author**: `https://github.com/sawyer-shi`
- **Email**: sawyer36@foxmail.com
- **License**: MIT License
- **Source Code**: `https://github.com/sawyer-shi/dify-plugins-file_converter`
- **Support**: Through Dify platform and GitHub Issues

---

**Ready to convert your files with ease?**



