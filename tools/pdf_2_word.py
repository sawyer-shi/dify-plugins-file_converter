import os
import tempfile
from collections.abc import Generator
from typing import Any, Dict, Optional
import json
import time
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    import fitz  # PyMuPDF
    from docx import Document
    from docx.shared import Inches
    PYPDF2_AVAILABLE = True
except ImportError:
    # Fallback for environments without required libraries
    PYPDF2_AVAILABLE = False

class PdfToWordTool(Tool):
    """Tool for converting PDF documents to Word format."""
    
    def get_file_info(self, file: File) -> dict:
        """
        获取文件信息
        Args:
            file: 文件对象
        Returns:
            文件信息字典
        """
        return {
            "filename": file.filename,
            "extension": file.extension,
            "mime_type": file.mime_type,
            "size": file.size,
            "url": file.url
        }
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # Get parameters
            file = tool_parameters.get("input_file")
            output_format = tool_parameters.get("output_format", "docx")
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
                
            # Validate input file format
            if not self._validate_input_file(file_info["extension"]):
                yield self.create_text_message("Error: Invalid file format. Only .pdf files are supported")
                return
                
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, 'wb') as f:
                    f.write(file.blob)
                
                # Update file info with the actual path
                file_info["path"] = input_path
                
                # Process conversion
                result = self._process_conversion(file_info, output_format, temp_dir)
                
                if result["success"]:
                    # Create output file info
                    output_files = []
                    for file_path in result["output_files"]:
                        output_files.append({
                            "filename": os.path.basename(file_path),
                            "size": os.path.getsize(file_path),
                            "path": file_path
                        })
                    
                    # Create JSON response
                    json_response = {
                        "success": True,
                        "conversion_type": "pdf_2_word",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"PDF converted to Word successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for file_path in result["output_files"]:
                        yield self.create_blob_message(blob=open(file_path, 'rb').read(), meta={"filename": os.path.basename(file_path)})
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _format_table(self, word_table, table_data):
        """Format the Word table with proper styling and preserve headers."""
        from docx.shared import Pt
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import nsmap
        
        # Fill the table with data and apply formatting
        for row_idx, row in enumerate(table_data):
            for col_idx, cell_data in enumerate(row):
                if cell_data is not None and str(cell_data).strip():
                    cell = word_table.cell(row_idx, col_idx)
                    cell.text = str(cell_data)
                    
                    # Format header row (first row) differently
                    if row_idx == 0:
                        # Make header text bold
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.bold = True
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Set header row background to light gray
                        try:
                            shading = cell._element.xpath('.//w:shd')[0]
                            shading.set('{%s}fill' % nsmap['w'], 'D9D9D9')
                        except:
                            pass  # Skip if shading fails
                    else:
                        # Align content to left for data rows
                        for paragraph in cell.paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Set table alignment
        word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Set column widths to be more proportional
        for row in word_table.rows:
            for cell in row.cells:
                cell.width = Inches(1.5)  # Set a reasonable default width
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for PDF to Word conversion."""
        return file_extension.lower() == ".pdf"
    
    def _process_conversion(self, file_info: Dict[str, Any], output_format: str, temp_dir: str) -> Dict[str, Any]:
        """Process the PDF to Word conversion using PyMuPDF library."""
        input_path = file_info["path"]
        output_files = []
        
        try:
            if not PYPDF2_AVAILABLE:
                return {"success": False, "message": "Required libraries (PyMuPDF, python-docx) are not available for PDF conversion"}
            
            # Default to docx if not specified
            if not output_format:
                output_format = "docx"
            elif output_format.lower() not in ["doc", "docx"]:
                output_format = "docx"
            
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.{output_format}")
            
            # Open the PDF file
            try:
                pdf_document = fitz.open(input_path)
            except Exception as e:
                return {"success": False, "message": f"Failed to open PDF file: {str(e)}"}
            
            # Create a new Word document
            doc = Document()
            
            # Process each page
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                
                # Add page heading
                doc.add_heading(f'Page {page_num + 1}', level=2)
                
                # Extract tables first
                tables = page.find_tables()
                if tables.tables:
                    for table_idx, table in enumerate(tables.tables):
                        try:
                            # Extract table data with bbox information
                            table_data = table.extract()
                            
                            # Create a table in Word document
                            if table_data and len(table_data) > 0:
                                # Determine the number of columns based on the first row
                                cols = len(table_data[0])
                                word_table = doc.add_table(rows=len(table_data), cols=cols)
                                word_table.style = 'Table Grid'
                                
                                # Apply formatting to the table
                                self._format_table(word_table, table_data)
                                
                                # Add a paragraph after the table
                                doc.add_paragraph(f"Table {table_idx + 1} from Page {page_num + 1}")
                                doc.add_paragraph()  # Add empty paragraph for spacing
                        except Exception as e:
                            # If table extraction fails, add a note
                            doc.add_paragraph(f"Note: Could not extract table {table_idx + 1} from Page {page_num + 1}: {str(e)}")
                            continue
                
                # Extract text (excluding tables)
                try:
                    # Get text blocks
                    text_dict = page.get_text("dict")
                    blocks = text_dict.get("blocks", [])
                    
                    page_text = ""
                    for block in blocks:
                        if "lines" in block:
                            for line in block["lines"]:
                                line_text = ""
                                for span in line.get("spans", []):
                                    line_text += span.get("text", "")
                                page_text += line_text + "\n"
                    
                    if page_text.strip():
                        doc.add_paragraph(page_text)
                except Exception as e:
                    # Fallback to simple text extraction if dict method fails
                    text = page.get_text()
                    if text.strip():
                        doc.add_paragraph(text)
                
                # Extract images
                image_list = page.get_images()
                for img_index, img in enumerate(image_list):
                    try:
                        # Get image data
                        xref = img[0]
                        pix = fitz.Pixmap(pdf_document, xref)
                        
                        # Skip CMYK images (not supported by python-docx)
                        if pix.n - pix.alpha < 4:
                            # Convert pixmap to bytes
                            img_data = pix.tobytes("png")
                            
                            # Create a temporary file to store the image
                            img_stream = io.BytesIO(img_data)
                            
                            # Add image to Word document
                            doc.add_picture(img_stream, width=Inches(6))
                            doc.add_paragraph(f"Image {img_index + 1} from Page {page_num + 1}")
                        
                        # Clean up
                        pix = None
                    except Exception as e:
                        # Log error but continue processing
                        print(f"Error processing image {img_index} on page {page_num + 1}: {str(e)}")
                        continue
                
                # Add page break except for the last page
                if page_num < len(pdf_document) - 1:
                    doc.add_page_break()
            
            # Close the PDF document
            pdf_document.close()
            
            # Save the Word document
            try:
                doc.save(output_path)
            except Exception as e:
                return {"success": False, "message": f"Failed to save Word document: {str(e)}"}
            
            # Check if file exists and has content
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                output_files.append(output_path)
            else:
                return {"success": False, "message": "Output file was not created or is empty"}
            
            return {
                "success": True, 
                "message": f"PDF converted to Word ({output_format}) successfully using PyMuPDF",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}