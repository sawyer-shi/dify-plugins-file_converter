import os
import tempfile
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    from PIL import Image
except ImportError:
    # Fallback for environments without PIL
    Image = None

class ImageToPdfTool(Tool):
    """Tool for converting image documents to PDF format."""
    
    def get_file_info(self, file: File) -> dict:
        """
        获取文件信息
        
        Args:
            file: 上传的文件对象
            
        Returns:
            dict: 文件信息
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
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
                
            # Validate input file format
            if not self._validate_input_file(file_info["extension"]):
                yield self.create_text_message("Error: Invalid file format. Only .jpg, .jpeg, .png, .bmp, and .tiff files are supported")
                return
                
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save the uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, "wb") as f:
                    f.write(file.blob)
                
                # Update file_info with the actual path
                file_info["path"] = input_path
                
                # Process conversion
                result = self._process_conversion(file_info, temp_dir)
                
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
                        "conversion_type": "image_2_pdf",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"Image converted to PDF successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for file_path in result["output_files"]:
                        yield self.create_blob_message(
                            blob=open(file_path, 'rb').read(), 
                            meta={
                                "filename": os.path.basename(file_path),
                                "mime_type": "application/pdf"
                            }
                        )
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for Image to PDF conversion."""
        return file_extension.lower() in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]
    
    def _process_conversion(self, file_info: Dict[str, Any], temp_dir: str) -> Dict[str, Any]:
        """Process the Image to PDF conversion."""
        input_path = file_info["path"]
        output_files = []
        
        try:
            if not Image:
                return {"success": False, "message": "PIL library is not available for Image conversion"}
            
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.pdf")
            
            # Convert Image to PDF using PIL
            image = Image.open(input_path)
            if image.mode == 'RGBA':
                image = image.convert('RGB')
            image.save(output_path, "PDF", resolution=100.0)
            output_files.append(output_path)
            
            return {
                "success": True, 
                "message": "Image converted to PDF successfully",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}