import os
import tempfile
from collections.abc import Generator
from typing import Any, Dict, Optional, List
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
    """Tool for converting multiple image documents to a single PDF format."""
    
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
            files = tool_parameters.get("input_files")
            
            if not files or len(files) == 0:
                yield self.create_text_message("Error: Missing required parameter 'input_files'")
                return
                
            # Get file info for all files
            files_info = []
            for file in files:
                file_info = self.get_file_info(file)
                
                # Validate input file format
                if not self._validate_input_file(file_info["extension"]):
                    yield self.create_text_message(f"Error: Invalid file format for {file_info['filename']}. Only .jpg, .jpeg, .png, .bmp, and .tiff files are supported")
                    return
                
                files_info.append(file_info)
                
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save all uploaded files to temp directory
                input_paths = []
                for i, file in enumerate(files):
                    file_info = files_info[i]
                    input_path = os.path.join(temp_dir, file_info["filename"])
                    with open(input_path, "wb") as f:
                        f.write(file.blob)
                    input_paths.append(input_path)
                
                # Process conversion
                result = self._process_conversion(input_paths, temp_dir)
                
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
                        "input_files": files_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"Images converted to PDF successfully: {result['message']}")
                    
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
    
    def _process_conversion(self, input_paths: List[str], temp_dir: str) -> Dict[str, Any]:
        """Process the multiple Images to PDF conversion."""
        output_files = []
        
        try:
            if not Image:
                return {"success": False, "message": "PIL library is not available for Image conversion"}
            
            # Generate output file path
            output_path = os.path.join(temp_dir, "combined_images.pdf")
            
            # Convert Images to PDF using PIL
            images = []
            for input_path in input_paths:
                image = Image.open(input_path)
                # Convert RGBA to RGB to avoid transparency issues in PDF
                if image.mode == 'RGBA':
                    image = image.convert('RGB')
                images.append(image)
            
            # Save all images as a single PDF
            if images:
                images[0].save(
                    output_path, 
                    "PDF", 
                    resolution=100.0,
                    save_all=True,
                    append_images=images[1:]
                )
                output_files.append(output_path)
            
            return {
                "success": True, 
                "message": f"Successfully converted {len(images)} images to PDF",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}