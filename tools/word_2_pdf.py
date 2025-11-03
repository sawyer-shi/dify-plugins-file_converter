import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    # Fallback for environments without docx2pdf
    DOCX2PDF_AVAILABLE = False

try:
    from docx import Document
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    DOCX_REPORTLAB_AVAILABLE = True
except ImportError:
    # Fallback for environments without docx and reportlab
    DOCX_REPORTLAB_AVAILABLE = False

class WordToPdfTool(Tool):
    """Tool for converting Word documents to PDF format."""
    
    def get_file_info(self, file: File) -> dict:
        """
        获取文件信息
        Args:
            file: 文件对象
        Returns:
            文件信息字典
        """
        file_info = {
            "filename": file.filename,
            "extension": file.extension,
            "mime_type": file.mime_type,
            "size": file.size,
            "url": file.url
        }
        
        # Add path attribute if it exists (for MockFile in testing)
        if hasattr(file, 'path'):
            file_info["path"] = file.path
            
        return file_info
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # Get input file parameter
            file = tool_parameters.get("input_file")
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
                
            # Validate input file format
            if not self._validate_input_file(file_info["extension"]):
                yield self.create_text_message("Error: Invalid file format. Only .doc and .docx files are supported")
                return
                
            # Create temporary directory for input file
            with tempfile.TemporaryDirectory() as temp_dir:
                # Use the test directory as output directory
                import os
                output_dir = r"D:\Work\Cursor\file_converter\test"
                os.makedirs(output_dir, exist_ok=True)
                
                # Save uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, 'wb') as f:
                    f.write(file.blob)
                
                # Update file info with the actual path
                file_info["path"] = input_path
                
                # Process conversion
                result = self._process_conversion(input_path, temp_dir)
                
                if result["success"]:
                    # Create output file info
                    output_files = []
                    for output_file_info in result["output_files"]:
                        output_files.append({
                            "filename": output_file_info["filename"],
                            "size": len(output_file_info["content"]),
                            "path": output_file_info["path"]
                        })
                    
                    # Create JSON response
                    json_response = {
                        "success": True,
                        "conversion_type": "word_2_pdf",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"Word document converted to PDF successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for output_file_info in result["output_files"]:
                        try:
                            # Use the pre-read content
                            if "content" in output_file_info:
                                yield self.create_blob_message(blob=output_file_info["content"], meta={"filename": output_file_info["filename"]})
                            else:
                                yield self.create_text_message(f"Error: No content available for file {output_file_info.get('filename', 'unknown')}")
                        except Exception as e:
                            yield self.create_text_message(f"Error sending file: {str(e)}")
                    
                    # Clean up only the temporary directory, not the output directory
                    try:
                        import shutil
                        # Only clean up the temp directory, not the output directory
                        # The temp directory will be automatically cleaned up by the context manager
                        pass
                    except Exception as e:
                        # Ignore cleanup errors
                        pass
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for Word to PDF conversion."""
        return file_extension.lower() in [".doc", ".docx"]
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """Process the Word to PDF conversion using docx2pdf library or fallback to docx+reportlab."""
        output_files = []
        
        # Generate output file path
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, f"{base_name}.pdf")
        
        # Try docx2pdf first if available
        if DOCX2PDF_AVAILABLE:
            # Try multiple conversion attempts
            conversion_success = False
            last_error = None
            
            for attempt in range(3):
                try:
                    # Convert Word to PDF using docx2pdf
                    convert(input_path, output_path)
                    
                    # Wait for file to be fully written and released
                    time.sleep(5)
                    
                    # Check if file exists and has content
                    if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                        conversion_success = True
                        break
                    else:
                        last_error = "Output file was not created or is empty"
                        
                except Exception as e:
                    last_error = str(e)
                    if attempt < 2:
                        time.sleep(3)  # Wait before retry
                        continue
            
            if conversion_success:
                # Try multiple times to read the file
                file_content = None
                for attempt in range(3):
                    try:
                        with open(output_path, 'rb') as f:
                            file_content = f.read()
                        break
                    except Exception as e:
                        if attempt < 2:
                            time.sleep(2)
                        else:
                            return {"success": False, "message": f"Error reading converted file: {str(e)}"}
                
                if file_content:
                    output_files.append({
                        "path": output_path,
                        "content": file_content,
                        "filename": f"{base_name}.pdf"
                    })
                    return {
                        "success": True,
                        "message": "Word document converted to PDF successfully using docx2pdf",
                        "output_files": output_files
                    }
        
        # Fallback to docx + reportlab if available
        if DOCX_REPORTLAB_AVAILABLE:
            try:
                # Load Word document
                doc = Document(input_path)
                
                # Create PDF document
                pdf_doc = SimpleDocTemplate(output_path, pagesize=A4)
                elements = []
                styles = getSampleStyleSheet()
                
                # Process each paragraph in the document
                for para in doc.paragraphs:
                    if para.text.strip():  # Skip empty paragraphs
                        p = Paragraph(para.text, styles['Normal'])
                        elements.append(p)
                        elements.append(Spacer(1, 0.1 * inch))
                
                # Process tables in the document
                for table in doc.tables:
                    from reportlab.platypus import Table, TableStyle
                    from reportlab.lib import colors
                    
                    # Extract table data
                    table_data = []
                    for row in table.rows:
                        row_data = []
                        for cell in row.cells:
                            row_data.append(cell.text)
                        table_data.append(row_data)
                    
                    # Create table with style
                    if table_data:
                        pdf_table = Table(table_data)
                        
                        # Add basic table style
                        style = TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 1), (-1, -1), 10),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ])
                        
                        pdf_table.setStyle(style)
                        elements.append(pdf_table)
                        elements.append(Spacer(1, 0.2 * inch))
                
                # Build PDF document
                pdf_doc.build(elements)
                
                # Wait for file to be fully written
                time.sleep(2)
                
                # Check if file exists and has content
                if not os.path.exists(output_path):
                    return {"success": False, "message": "Output PDF file was not created using docx+reportlab"}
                    
                if os.path.getsize(output_path) == 0:
                    return {"success": False, "message": "Output PDF file is empty using docx+reportlab"}
                
                # Try multiple times to read the file
                file_content = None
                for attempt in range(3):
                    try:
                        with open(output_path, 'rb') as f:
                            file_content = f.read()
                        break
                    except Exception as e:
                        if attempt < 2:
                            time.sleep(2)
                        else:
                            return {"success": False, "message": f"Error reading converted file: {str(e)}"}
                
                if file_content:
                    output_files.append({
                        "path": output_path,
                        "content": file_content,
                        "filename": f"{base_name}.pdf"
                    })
                    return {
                        "success": True, 
                        "message": "Word document converted to PDF successfully using docx+reportlab",
                        "output_files": output_files
                    }
                else:
                    return {"success": False, "message": "Failed to read converted file after multiple attempts using docx+reportlab"}
                    
            except Exception as e:
                return {"success": False, "message": f"Error converting with docx+reportlab: {str(e)}"}
        
        # If we reach here, neither method worked
        if not DOCX2PDF_AVAILABLE and not DOCX_REPORTLAB_AVAILABLE:
            return {"success": False, "message": "Neither docx2pdf nor docx+reportlab libraries are available for Word conversion"}
        else:
            return {"success": False, "message": f"Failed to convert Word document using available methods"}