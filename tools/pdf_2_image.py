import os
import tempfile
from collections.abc import Generator
from typing import Any, Dict, Optional
import json
import time

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    from pdf2image import convert_from_path
    from PIL import Image
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    # Fallback for environments without pdf2image
    PDF2IMAGE_AVAILABLE = False

class PdfToImageTool(Tool):
    """Tool for converting PDF documents to image format."""
    
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
            output_format = tool_parameters.get("output_format", "png")
            
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
                # Save the uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, "wb") as f:
                    f.write(file.blob)
                
                # Update file_info with the actual path
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
                        "conversion_type": "pdf_2_image",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"PDF converted to images successfully: {result['message']}")
                    
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
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for PDF to Image conversion."""
        return file_extension.lower() == ".pdf"
    
    def _process_conversion(self, file_info: Dict[str, Any], output_format: str, temp_dir: str) -> Dict[str, Any]:
        """Process the PDF to Image conversion using pdf2image library."""
        input_path = file_info["path"]
        output_files = []
        
        try:
            if not PDF2IMAGE_AVAILABLE:
                return {"success": False, "message": "pdf2image library is not available for PDF conversion"}
            
            # Default to png if not specified
            if not output_format:
                output_format = "png"
            elif output_format.lower() not in ["jpg", "jpeg", "png", "bmp", "tiff"]:
                output_format = "png"
            
            # Set poppler path
            poppler_path = r"D:\Work\Cursor\file_converter\poppler\poppler-23.07.0\Library\bin"
            
            # Convert PDF to images using pdf2image
            # dpi=300 for good quality, fmt=output_format for the desired format
            images = convert_from_path(input_path, dpi=300, fmt=output_format.lower(), poppler_path=poppler_path)
            
            if not images:
                return {"success": False, "message": "No images were generated"}
            
            # Save each image to a file
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            for i, image in enumerate(images):
                output_filename = f"{base_name}_{i+1:03d}.{output_format.lower()}"
                output_path = os.path.join(temp_dir, output_filename)
                
                # Save the image
                if output_format.lower() == "jpg" or output_format.lower() == "jpeg":
                    # JPEG doesn't support transparency
                    if image.mode in ("RGBA", "LA", "P"):
                        # Convert to RGB mode for JPEG
                        background = Image.new("RGB", image.size, (255, 255, 255))
                        if image.mode == "P":
                            image = image.convert("RGBA")
                        background.paste(image, mask=image.split()[-1] if image.mode == "RGBA" else None)
                        image = background
                    image.save(output_path, "JPEG", quality=95)
                else:
                    image.save(output_path, output_format.upper())
                
                output_files.append(output_path)
            
            return {
                "success": True, 
                "message": f"PDF converted to {len(output_files)} {output_format} images successfully",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}