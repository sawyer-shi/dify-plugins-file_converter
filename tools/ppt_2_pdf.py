import os
import tempfile
import time
import subprocess
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

class PptToPdfTool(Tool):
    """Tool for converting PowerPoint documents to PDF format."""
    
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
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
                
            # Validate input file format
            if not self._validate_input_file(file_info["extension"]):
                yield self.create_text_message("Error: Invalid file format. Only .ppt and .pptx files are supported")
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
                result = self._process_conversion(file_info, temp_dir)
                
                if result["success"]:
                    # Create output file info
                    output_files = []
                    for file_info in result["output_files"]:
                        output_files.append({
                            "filename": file_info["filename"],
                            "size": len(file_info["content"]),
                            "path": file_info["path"]
                        })
                    
                    # Create JSON response
                    json_response = {
                        "success": True,
                        "conversion_type": "ppt_2_pdf",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"PPT converted to PDF successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for file_info in result["output_files"]:
                        try:
                            # Use the pre-read content
                            if "content" in file_info:
                                yield self.create_blob_message(blob=file_info["content"], meta={"filename": file_info["filename"]})
                            else:
                                yield self.create_text_message(f"Error: No content available for file {file_info.get('filename', 'unknown')}")
                        except Exception as e:
                            yield self.create_text_message(f"Error sending file: {str(e)}")
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for PPT to PDF conversion."""
        return file_extension.lower() in [".ppt", ".pptx"]
    
    def _process_conversion(self, file_info: Dict[str, Any], temp_dir: str) -> Dict[str, Any]:
        """Process the PowerPoint to PDF conversion."""
        input_path = file_info["path"]
        output_files = []
        
        try:
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.pdf")
            
            # Try different methods for conversion
            conversion_success = False
            
            # Method 1: Try using LibreOffice (if available)
            try:
                # Command to convert PPT to PDF using LibreOffice
                # --headless: run without GUI
                # --convert-to pdf: convert to PDF format
                # --outdir: output directory
                cmd = [
                    "soffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", temp_dir,
                    input_path
                ]
                
                # Run the command
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                
                if result.returncode == 0:
                    # Check if the output file was created
                    if os.path.exists(output_path):
                        conversion_success = True
                else:
                    print(f"LibreOffice conversion failed: {result.stderr}")
            except Exception as e:
                print(f"LibreOffice not available or failed: {str(e)}")
            
            # Method 2: Try using unoconv (if available)
            if not conversion_success:
                try:
                    cmd = ["unoconv", "-f", "pdf", "-o", output_path, input_path]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                    
                    if result.returncode == 0 and os.path.exists(output_path):
                        conversion_success = True
                    else:
                        print(f"unoconv conversion failed: {result.stderr}")
                except Exception as e:
                    print(f"unoconv not available or failed: {str(e)}")
            
            # Method 3: Try using Aspose.Slides (if available)
            if not conversion_success:
                try:
                    import aspose.slides as slides
                    import aspose.pdf as pdf
                    
                    # Load the PowerPoint presentation
                    presentation = slides.Presentation(input_path)
                    
                    # Save as PDF
                    presentation.save(output_path, slides.export.SaveFormat.PDF)
                    
                    if os.path.exists(output_path):
                        conversion_success = True
                except Exception as e:
                    print(f"Aspose.Slides conversion failed: {str(e)}")
            
            # Method 4: Try using a more comprehensive Python-PPTX and ReportLab approach
            if not conversion_success:
                try:
                    from pptx import Presentation
                    from pptx.enum.shapes import MSO_SHAPE_TYPE
                    from reportlab.pdfgen import canvas
                    from reportlab.lib.pagesizes import letter, A4
                    from reportlab.lib.utils import ImageReader
                    from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
                    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                    from reportlab.lib import colors
                    from reportlab.lib.units import inch
                    import io
                    
                    # Open the PowerPoint presentation
                    prs = Presentation(input_path)
                    
                    # Create PDF with A4 pagesize
                    c = canvas.Canvas(output_path, pagesize=A4)
                    width, height = A4
                    
                    # Define margins
                    left_margin = 0.5 * inch
                    right_margin = 0.5 * inch
                    top_margin = 0.5 * inch
                    bottom_margin = 0.5 * inch
                    content_width = width - left_margin - right_margin
                    content_height = height - top_margin - bottom_margin
                    
                    # Process each slide
                    for slide_idx, slide in enumerate(prs.slides):
                        # Add a new page for each slide (except the first one)
                        if slide_idx > 0:
                            c.showPage()
                        
                        # Track current y position
                        current_y = height - top_margin
                        
                        # Process shapes in order
                        for shape in slide.shapes:
                            # Skip placeholder shapes
                            if shape.is_placeholder:
                                continue
                                
                            # Text frames
                            if shape.has_text_frame:
                                try:
                                    text = shape.text_frame.text
                                    if text.strip():
                                        # Add text to PDF
                                        c.setFont("Helvetica", 12)
                                        for line in text.split('\n'):
                                            if current_y < bottom_margin + 20:  # Check if we need a new page
                                                c.showPage()
                                                current_y = height - top_margin
                                            
                                            c.drawString(left_margin, current_y, line)
                                            current_y -= 20
                                except Exception as e:
                                    print(f"Error processing text shape: {str(e)}")
                                    continue
                            
                            # Tables
                            elif shape.has_table:
                                try:
                                    table = shape.table
                                    table_data = []
                                    
                                    # Extract table data
                                    for row_idx, row in enumerate(table.rows):
                                        row_data = []
                                        for cell_idx, cell in enumerate(row.cells):
                                            row_data.append(cell.text_frame.text)
                                        table_data.append(row_data)
                                    
                                    # Create a ReportLab table
                                    if table_data:
                                        # Calculate column widths
                                        cols = len(table_data[0])
                                        col_widths = [content_width / cols] * cols
                                        
                                        # Create table
                                        rl_table = Table(table_data, colWidths=col_widths)
                                        rl_table.setStyle(TableStyle([
                                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                                        ]))
                                        
                                        # Draw the table
                                        table_width, table_height = rl_table.wrapOn(c, content_width, content_height)
                                        
                                        if current_y - table_height < bottom_margin:  # Check if we need a new page
                                            c.showPage()
                                            current_y = height - top_margin
                                        
                                        rl_table.drawOn(c, left_margin, current_y - table_height)
                                        current_y -= table_height + 20
                                except Exception as e:
                                    print(f"Error processing table: {str(e)}")
                                    continue
                            
                            # Images
                            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                try:
                                    # Get image data
                                    image_bytes = shape.image.blob
                                    
                                    # Create a temporary image file
                                    img_stream = io.BytesIO(image_bytes)
                                    img_reader = ImageReader(img_stream)
                                    
                                    # Get image dimensions
                                    img_width, img_height = img_reader.getSize()
                                    
                                    # Calculate scaling to fit within content area
                                    scale = min(content_width / img_width, content_height / img_height, 1.0)
                                    scaled_width = img_width * scale
                                    scaled_height = img_height * scale
                                    
                                    if current_y - scaled_height < bottom_margin:  # Check if we need a new page
                                        c.showPage()
                                        current_y = height - top_margin
                                    
                                    # Draw the image
                                    c.drawImage(img_reader, left_margin, current_y - scaled_height, 
                                               width=scaled_width, height=scaled_height)
                                    current_y -= scaled_height + 20
                                except Exception as e:
                                    print(f"Error processing image: {str(e)}")
                                    continue
                    
                    # Save the PDF
                    c.save()
                    
                    if os.path.exists(output_path):
                        conversion_success = True
                except Exception as e:
                    print(f"Enhanced Python-PPTX conversion failed: {str(e)}")
                except Exception as e:
                    print(f"Python-PPTX conversion failed: {str(e)}")
            
            if not conversion_success:
                return {"success": False, "message": "All conversion methods failed. Please install LibreOffice, unoconv, or python-pptx with reportlab."}
            
            # Wait for file to be fully written
            time.sleep(2)
            
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
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts"}
            
            return {
                "success": True, 
                "message": "PowerPoint presentation converted to PDF successfully",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}