import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# Try to import PyMuPDF for PDF text extraction
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Try to import pdfplumber as alternative for PDF text extraction
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

class PdfToTextTool(Tool):
    """Tool for converting PDF files to text format."""
    
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
            file_info["path"] = file_info
            
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
            if not self._validate_input_file(file_info):
                yield self.create_text_message("Error: Invalid file format. Only PDF files (.pdf) are supported")
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
                        "conversion_type": "pdf_2_text",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"PDF file converted to text successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for file_info in result["output_files"]:
                        try:
                            # Use the pre-read content
                            if "content" in file_info:
                                yield self.create_blob_message(
                                    blob=file_info["content"], 
                                    meta={
                                        "filename": file_info["filename"],
                                        "mime_type": "text/plain"
                                    }
                                )
                            else:
                                yield self.create_text_message(f"Error: No content available for file {file_info.get('filename', 'unknown')}")
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
    
    def _validate_input_file(self, file_info: dict) -> bool:
        """Validate if the input file is a valid PDF file."""
        # Check file extension
        if not file_info["extension"].lower().endswith('.pdf'):
            return False
            
        # Check if file is readable as PDF
        if "path" in file_info:
            try:
                if PYMUPDF_AVAILABLE:
                    doc = fitz.open(file_info["path"])
                    doc.close()
                    return True
                elif PDFPLUMBER_AVAILABLE:
                    with pdfplumber.open(file_info["path"]) as pdf:
                        _ = pdf.pages[0]  # Try to access first page
                    return True
                else:
                    # If no PDF library is available, just check file extension
                    return True
            except Exception:
                return False
        
        # If path not available, just check file extension
        return True
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """Process the PDF to text conversion using PyMuPDF or pdfplumber."""
        output_files = []
        
        # Generate output file path
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, f"{base_name}.txt")
        
        # Check if required libraries are available
        if not PYMUPDF_AVAILABLE and not PDFPLUMBER_AVAILABLE:
            return {"success": False, "message": "Required library (PyMuPDF or pdfplumber) is not available. Please install one using: pip install PyMuPDF or pip install pdfplumber"}
        
        try:
            text_content = ""
            
            # Try PyMuPDF first (faster and more reliable)
            if PYMUPDF_AVAILABLE:
                try:
                    doc = fitz.open(input_path)
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        text_content += page.get_text()
                        if page_num < len(doc) - 1:
                            text_content += "\n\n--- Page " + str(page_num + 1) + " ---\n\n"
                    doc.close()
                except Exception as e:
                    # If PyMuPDF fails, try pdfplumber
                    if PDFPLUMBER_AVAILABLE:
                        text_content = self._extract_with_pdfplumber(input_path)
                    else:
                        return {"success": False, "message": f"Error extracting text with PyMuPDF: {str(e)}"}
            else:
                # Use pdfplumber if PyMuPDF is not available
                text_content = self._extract_with_pdfplumber(input_path)
            
            # Write text to file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text_content)
            
            # Wait for file to be fully written
            time.sleep(1)
            
            # Check if file exists and has content
            if not os.path.exists(output_path):
                return {"success": False, "message": "Output text file was not created"}
                
            if os.path.getsize(output_path) == 0:
                return {"success": False, "message": "Output text file is empty"}
            
            # Try multiple times to read the file
            file_content = None
            for attempt in range(3):
                try:
                    with open(output_path, 'r', encoding='utf-8') as f:
                        file_content = f.read()
                    break
                except Exception as e:
                    if attempt < 2:
                        time.sleep(1)
                    else:
                        return {"success": False, "message": f"Error reading converted file: {str(e)}"}
            
            if file_content:
                output_files.append({
                    "path": output_path,
                    "content": file_content.encode('utf-8'),
                    "filename": f"{base_name}.txt"
                })
                library_used = "PyMuPDF" if PYMUPDF_AVAILABLE else "pdfplumber"
                return {
                    "success": True, 
                    "message": f"PDF file converted to text successfully using {library_used}",
                    "output_files": output_files
                }
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts"}
                    
        except Exception as e:
            return {"success": False, "message": f"Error converting PDF to text: {str(e)}"}
    
    def _extract_with_pdfplumber(self, input_path: str) -> str:
        """Extract text from PDF using pdfplumber."""
        text_content = ""
        try:
            with pdfplumber.open(input_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        text_content += page_text
                        if page_num < len(pdf.pages) - 1:
                            text_content += "\n\n--- Page " + str(page_num + 1) + " ---\n\n"
        except Exception as e:
            raise Exception(f"Error extracting text with pdfplumber: {str(e)}")
        
        return text_content