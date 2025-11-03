import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional
import json
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# Try to import reportlab components for Chinese font support
try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.fonts import addMapping
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Try to import python-pptx and reportlab components for conversion
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
    PPTX_REPORTLAB_AVAILABLE = True
except ImportError:
    PPTX_REPORTLAB_AVAILABLE = False

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
    
    def _register_chinese_fonts(self):
        """Register Chinese fonts for reportlab to use."""
        if not REPORTLAB_AVAILABLE:
            return
            
        try:
            # Get the directory of the current script
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # Get the fonts directory (one level up from tools directory)
            fonts_dir = os.path.join(os.path.dirname(current_dir), "fonts")
            
            # Try to register common Chinese fonts available on Windows
            font_paths = [
                # Project Chinese font (highest priority)
                ('ChineseFont', os.path.join(fonts_dir, "chinese_font.ttc")),
                # SimSun (宋体)
                ('SimSun', 'C:/Windows/Fonts/simsun.ttc'),
                ('SimSun', 'C:/Windows/Fonts/simsun.ttf'),
                # SimHei (黑体)
                ('SimHei', 'C:/Windows/Fonts/simhei.ttf'),
                # Microsoft YaHei (微软雅黑)
                ('Microsoft YaHei', 'C:/Windows/Fonts/msyh.ttf'),
                # KaiTi (楷体)
                ('KaiTi', 'C:/Windows/Fonts/kaiti.ttf'),
                # FangSong (仿宋)
                ('FangSong', 'C:/Windows/Fonts/simfang.ttf'),
            ]
            
            registered_fonts = []
            for font_name, font_path in font_paths:
                try:
                    if os.path.exists(font_path):
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        registered_fonts.append(font_name)
                        
                        # If this is the ChineseFont, also register ChineseFont-Bold
                        if font_name == "ChineseFont":
                            pdfmetrics.registerFont(TTFont("ChineseFont-Bold", font_path))
                            registered_fonts.append("ChineseFont-Bold")
                except Exception as e:
                    # Continue trying other fonts if one fails
                    continue
            
            # Register bold variants if available
            bold_variants = [
                ('SimSun-Bold', 'C:/Windows/Fonts/simsunb.ttf'),
                ('SimHei-Bold', 'C:/Windows/Fonts/simheib.ttf'),
            ]
            
            for font_name, font_path in bold_variants:
                try:
                    if os.path.exists(font_path):
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        registered_fonts.append(font_name)
                except Exception as e:
                    # Continue trying other fonts if one fails
                    continue
            
            # If no Chinese fonts were registered, create a fallback mapping
            if not registered_fonts:
                # Map Chinese font names to available fonts as fallback
                font_mapping = {
                    'ChineseFont': 'Helvetica',
                    'ChineseFont-Bold': 'Helvetica-Bold',
                    'SimSun': 'Helvetica',
                    'SimHei': 'Helvetica',
                    'SimSun-Bold': 'Helvetica-Bold',
                    'SimHei-Bold': 'Helvetica-Bold',
                    'Microsoft YaHei': 'Helvetica',
                    'KaiTi': 'Helvetica',
                    'FangSong': 'Helvetica',
                }
                
                for chinese_font, fallback_font in font_mapping.items():
                    try:
                        # Create an alias for the fallback font
                        if 'Bold' in chinese_font and 'Bold' in fallback_font:
                            addMapping(chinese_font, 0, 0, fallback_font)
                        else:
                            addMapping(chinese_font, 0, 0, fallback_font)
                    except Exception:
                        # If even fallback fails, just continue
                        continue
                        
        except Exception as e:
            # If font registration fails completely, we'll rely on default fonts
            pass
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # Get parameters
            file = tool_parameters.get("input_file")
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
                
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, 'wb') as f:
                    f.write(file.blob)
                
                # Update file info with the actual path
                file_info["path"] = input_path
                
                # Validate input file format
                if not self._validate_input_file(file_info):
                    yield self.create_text_message("Error: Invalid file format or missing dependencies. Only .ppt and .pptx files are supported, and python-pptx with reportlab is required.".encode('utf-8', errors='replace').decode('utf-8'))
                    return
                    
                # Process conversion
                result = self._process_conversion(file_info, temp_dir)
                
                if result["success"]:
                    # Create output file info
                    output_files = []
                    for file_info in result["output_files"]:
                        # Create safe file info with UTF-8 encoding
                        safe_file_info = {
                            "filename": file_info["filename"].encode('utf-8', errors='replace').decode('utf-8'),
                            "size": len(file_info["content"]),
                            "path": file_info["path"].encode('utf-8', errors='replace').decode('utf-8')
                        }
                        output_files.append(safe_file_info)
                    
                    # Create a safe copy of file_info for JSON serialization
                    safe_file_info = {}
                    for key, value in file_info.items():
                        if isinstance(value, str):
                            # Ensure string values are valid UTF-8
                            safe_file_info[key] = value.encode('utf-8', errors='replace').decode('utf-8')
                        else:
                            safe_file_info[key] = value
                    
                    # Create JSON response with safe data
                    json_response = {
                        "success": True,
                        "conversion_type": "ppt_2_pdf",
                        "input_file": safe_file_info,
                        "output_files": output_files,
                        "message": result["message"].encode('utf-8', errors='replace').decode('utf-8')
                    }
                    
                    # Ensure all string fields in json_response are properly UTF-8 encoded
                    for key, value in json_response.items():
                        if isinstance(value, str):
                            json_response[key] = value.encode('utf-8', errors='replace').decode('utf-8')
                        elif isinstance(value, dict):
                            for sub_key, sub_value in value.items():
                                if isinstance(sub_value, str):
                                    json_response[key][sub_key] = sub_value.encode('utf-8', errors='replace').decode('utf-8')
                        elif isinstance(value, list):
                            for i, item in enumerate(value):
                                if isinstance(item, dict):
                                    for sub_key, sub_value in item.items():
                                        if isinstance(sub_value, str):
                                            json_response[key][i][sub_key] = sub_value.encode('utf-8', errors='replace').decode('utf-8')
                    
                    # Send text message
                    yield self.create_text_message(f"PPT converted to PDF successfully: {result['message']}".encode('utf-8', errors='replace').decode('utf-8'))
                    
                    # Send JSON message
                    json_response = {
                        "success": True,
                        "message": "PPT converted to PDF successfully".encode('utf-8', errors='replace').decode('utf-8'),
                        "output_files": output_files
                    }
                    
                    # Ensure all string values in json_response are properly UTF-8 encoded
                    for key, value in json_response.items():
                        if isinstance(value, str):
                            json_response[key] = value.encode('utf-8', errors='replace').decode('utf-8')
                        elif isinstance(value, dict):
                            for sub_key, sub_value in value.items():
                                if isinstance(sub_value, str):
                                    json_response[key][sub_key] = sub_value.encode('utf-8', errors='replace').decode('utf-8')
                        elif isinstance(value, list):
                            for i, item in enumerate(value):
                                if isinstance(item, dict):
                                    for sub_key, sub_value in item.items():
                                        if isinstance(sub_value, str):
                                            json_response[key][i][sub_key] = sub_value.encode('utf-8', errors='replace').decode('utf-8')
                    
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
                                        "mime_type": "application/pdf"
                                    }
                                )
                            else:
                                yield self.create_text_message(f"Error: No content available for file {file_info.get('filename', 'unknown')}")
                        except Exception as e:
                            yield self.create_text_message(f"Error sending file: {str(e)}")
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}".encode('utf-8', errors='replace').decode('utf-8'))
                    
        except Exception as e:
            yield self.create_text_message(f"Error during PPT to PDF conversion: {str(e)}".encode('utf-8', errors='replace').decode('utf-8'))
    
    def _validate_input_file(self, file_info: dict) -> bool:
        """Validate if the input file format is supported for PPT to PDF conversion."""
        # Check file extension
        if not file_info["extension"].lower().endswith(('.ppt', '.pptx')):
            return False
            
        # Check if python-pptx and reportlab are available
        if not PPTX_REPORTLAB_AVAILABLE:
            return False
            
        # Try to load the file with python-pptx to verify it's a valid PPT file
        if "path" in file_info:
            try:
                prs = Presentation(file_info["path"])
                # Just access the slides to verify the file is readable
                _ = list(prs.slides)
                return True
            except Exception as e:
                return False
        
        # If path not available, just check file extension and dependencies
        return True
    
    def _process_conversion(self, file_info: Dict[str, Any], temp_dir: str) -> Dict[str, Any]:
        """Process the PowerPoint to PDF conversion using python-pptx and reportlab."""
        input_path = file_info["path"]
        output_files = []
        
        try:
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.pdf")
            
            # Check if python-pptx and reportlab are available
            if not PPTX_REPORTLAB_AVAILABLE:
                return {"success": False, "message": "Required libraries (python-pptx and reportlab) are not available. Please install them to use this tool.".encode('utf-8', errors='replace').decode('utf-8')}
            
            # Register Chinese fonts
            self._register_chinese_fonts()
            
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
                                # Use Chinese font for text
                                try:
                                    # Try to use ChineseFont first, then fallback to other fonts
                                    if "ChineseFont" in pdfmetrics.getRegisteredFontNames():
                                        c.setFont("ChineseFont", 12)
                                    elif "SimSun" in pdfmetrics.getRegisteredFontNames():
                                        c.setFont("SimSun", 12)  # Use SimSun for Chinese text
                                    else:
                                        c.setFont("Helvetica", 12)  # Fallback to Helvetica
                                except:
                                    c.setFont("Helvetica", 12)  # Fallback to Helvetica
                                
                                # Process text line by line, handling potential encoding issues
                                for line in text.split('\n'):
                                    if current_y < bottom_margin + 20:  # Check if we need a new page
                                        c.showPage()
                                        current_y = height - top_margin
                                    
                                    # Ensure text is properly encoded
                                    try:
                                        # Try to encode as UTF-8 and decode to handle any encoding issues
                                        safe_line = line.encode('utf-8', errors='replace').decode('utf-8')
                                        c.drawString(left_margin, current_y, safe_line)
                                    except Exception:
                                        # If encoding still fails, use a safe representation
                                        safe_line = str(line).encode('utf-8', errors='replace').decode('utf-8')
                                        c.drawString(left_margin, current_y, safe_line)
                                    
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
                                    # Safely extract text from cell
                                    try:
                                        cell_text = cell.text_frame.text
                                        # Ensure text is properly encoded
                                        safe_text = cell_text.encode('utf-8', errors='replace').decode('utf-8')
                                        row_data.append(safe_text)
                                    except Exception:
                                        row_data.append("")  # Use empty string if extraction fails
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
                                    ('FONTNAME', (0, 0), (-1, 0), 'ChineseFont-Bold'),  # Use Chinese font for header
                                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                    ('FONTNAME', (0, 1), (-1, -1), 'ChineseFont'),  # Use Chinese font for cells
                                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
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
                        return {"success": False, "message": f"Error reading converted file: {str(e)}".encode('utf-8', errors='replace').decode('utf-8')}
            
            if file_content:
                # Ensure filename and path are properly UTF-8 encoded
                safe_filename = f"{base_name}.pdf".encode('utf-8', errors='replace').decode('utf-8')
                safe_path = output_path.encode('utf-8', errors='replace').decode('utf-8')
                
                output_files.append({
                    "path": safe_path,
                    "content": file_content,
                    "filename": safe_filename
                })
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts".encode('utf-8', errors='replace').decode('utf-8')}
            
            return {
                "success": True, 
                "message": "PowerPoint presentation converted to PDF successfully using python-pptx and reportlab".encode('utf-8', errors='replace').decode('utf-8'),
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}".encode('utf-8', errors='replace').decode('utf-8')}