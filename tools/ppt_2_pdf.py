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
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
    from pptx.enum.dml import MSO_THEME_COLOR_INDEX
    from pptx.util import Inches, Emu
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    PPTX_REPORTLAB_AVAILABLE = True
except ImportError:
    PPTX_REPORTLAB_AVAILABLE = False

class PptToPdfTool(Tool):
    """改进的PPT转PDF工具，解决排版、隐藏元素和页数问题"""
    
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
    
    def _is_hidden_shape(self, shape) -> bool:
        """
        检查形状是否为隐藏元素
        Args:
            shape: PPT形状对象
        Returns:
            bool: 是否为隐藏元素
        """
        try:
            # 检查形状是否可见
            if hasattr(shape, 'visible') and not shape.visible:
                return True
                
            # 检查形状是否为隐藏的占位符
            if shape.is_placeholder:
                # 检查占位符类型，跳过页眉页脚等隐藏占位符
                placeholder_format = shape.placeholder_format
                if placeholder_format.type in [1, 2, 3, 10, 11, 12, 13, 14, 15]:  # 常见的隐藏占位符类型
                    return True
            
            # 检查形状是否在幻灯片外
            if hasattr(shape, 'left') and hasattr(shape, 'top') and hasattr(shape, 'width') and hasattr(shape, 'height'):
                slide_width = shape.slide.slide_dimensions.width
                slide_height = shape.slide.slide_dimensions.height
                
                # 如果形状完全在幻灯片外，则视为隐藏
                if (shape.left > slide_width or 
                    shape.top > slide_height or
                    shape.left + shape.width < 0 or
                    shape.top + shape.height < 0):
                    return True
                    
            # 检查形状透明度
            if hasattr(shape, 'fill') and hasattr(shape.fill, 'transparency'):
                if shape.fill.transparency > 0.8:  # 如果透明度很高，可能是隐藏元素
                    return True
                    
            # 检查形状是否为隐藏的图标或装饰性元素
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                if hasattr(shape, 'auto_shape_type'):
                    # 跳过常见的装饰性形状
                    if shape.auto_shape_type in [
                        MSO_AUTO_SHAPE_TYPE.OVAL, 
                        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
                        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                        MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
                        MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE,
                        MSO_AUTO_SHAPE_TYPE.PARALLELOGRAM,
                        MSO_AUTO_SHAPE_TYPE.TRAPEZOID,
                        MSO_AUTO_SHAPE_TYPE.DIAMOND,
                        MSO_AUTO_SHAPE_TYPE.PENTAGON,
                        MSO_AUTO_SHAPE_TYPE.HEXAGON,
                        MSO_AUTO_SHAPE_TYPE.OCTAGON,
                        MSO_AUTO_SHAPE_TYPE.STAR_10_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_12_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_16_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_24_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_32_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_4_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_5_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_6_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_7_POINT,
                        MSO_AUTO_SHAPE_TYPE.STAR_8_POINT
                    ]:
                        # 检查形状是否很小（可能是装饰性图标）
                        if shape.width < 0.5 * inch and shape.height < 0.5 * inch:
                            return True
                            
            return False
        except Exception:
            # 如果检查过程中出现异常，默认不隐藏
            return False
    
    def _get_shape_position(self, shape) -> Dict[str, float]:
        """
        获取形状的位置和大小信息
        Args:
            shape: PPT形状对象
        Returns:
            Dict: 包含位置和大小信息的字典
        """
        try:
            return {
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            }
        except Exception:
            # 如果获取失败，返回默认值
            return {
                'left': 0,
                'top': 0,
                'width': 100,
                'height': 100
            }
    
    def _get_text_formatting(self, shape) -> Dict[str, Any]:
        """
        获取文本格式信息
        Args:
            shape: PPT形状对象
        Returns:
            Dict: 包含文本格式信息的字典
        """
        try:
            formatting = {
                'font_name': 'Helvetica',
                'font_size': 12,
                'bold': False,
                'italic': False,
                'color': (0, 0, 0),  # 默认黑色
                'alignment': TA_LEFT
            }
            
            if shape.has_text_frame and shape.text_frame.paragraphs:
                # 获取第一个段落的格式作为代表
                paragraph = shape.text_frame.paragraphs[0]
                if paragraph.runs:
                    run = paragraph.runs[0]
                    
                    # 获取字体信息
                    if hasattr(run, 'font'):
                        if run.font.name:
                            formatting['font_name'] = run.font.name
                        if run.font.size:
                            formatting['font_size'] = run.font.size.pt
                        if run.font.bold:
                            formatting['bold'] = run.font.bold
                        if run.font.italic:
                            formatting['italic'] = run.font.italic
                        
                        # 获取字体颜色
                        if run.font.color and run.font.color.type == MSO_THEME_COLOR_INDEX.RGB:
                            formatting['color'] = (
                                run.font.color.rgb.red,
                                run.font.color.rgb.green,
                                run.font.color.rgb.blue
                            )
                    
                    # 获取对齐方式
                    if hasattr(paragraph, 'alignment') and paragraph.alignment:
                        alignment_map = {
                            0: TA_LEFT,
                            1: TA_CENTER,
                            2: TA_RIGHT,
                            3: TA_JUSTIFY
                        }
                        formatting['alignment'] = alignment_map.get(paragraph.alignment, TA_LEFT)
            
            return formatting
        except Exception:
            # 如果获取失败，返回默认格式
            return {
                'font_name': 'Helvetica',
                'font_size': 12,
                'bold': False,
                'italic': False,
                'color': (0, 0, 0),
                'alignment': TA_LEFT
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
                    
                    # Send text message with conversion details
                    yield self.create_text_message(f"PPT converted to PDF successfully: {result['message']}".encode('utf-8', errors='replace').decode('utf-8'))
                    
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
        """Process the PowerPoint to PDF conversion using python-pptx and reportlab with improved layout preservation."""
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
            
            # Get slide dimensions to maintain aspect ratio
            # For pptx, we need to get the dimensions from the presentation
            slide_width = prs.slide_width
            slide_height = prs.slide_height
            
            # Create PDF with the same aspect ratio as the PPT
            # Convert from EMUs to points (1 inch = 72 points = 914400 EMUs)
            pdf_width = slide_width / 12700  # Convert EMUs to points
            pdf_height = slide_height / 12700  # Convert EMUs to points
            
            # Create PDF with custom pagesize matching PPT dimensions
            c = canvas.Canvas(output_path, pagesize=(pdf_width, pdf_height))
            
            # Process each slide
            for slide_idx, slide in enumerate(prs.slides):
                # Add a new page for each slide (except the first one)
                if slide_idx > 0:
                    c.showPage()
                
                # Process shapes in order, maintaining their original positions
                # First, collect all non-hidden shapes
                visible_shapes = []
                for shape in slide.shapes:
                    # Skip hidden shapes
                    if self._is_hidden_shape(shape):
                        continue
                    visible_shapes.append(shape)
                
                # Process shapes by type to maintain proper layering
                # Background first, then shapes, then text on top
                background_shapes = []
                image_shapes = []
                table_shapes = []
                text_shapes = []
                
                for shape in visible_shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        image_shapes.append(shape)
                    elif shape.has_table:
                        table_shapes.append(shape)
                    elif shape.has_text_frame:
                        text_shapes.append(shape)
                    else:
                        background_shapes.append(shape)
                
                # Process shapes in order: background, images, tables, text
                for shape_list in [background_shapes, image_shapes, table_shapes, text_shapes]:
                    for shape in shape_list:
                        try:
                            # Get shape position
                            pos = self._get_shape_position(shape)
                            
                            # Convert PPT coordinates to PDF coordinates
                            # PPT origin is top-left, PDF origin is bottom-left
                            x = pos['left'] / 12700  # Convert EMUs to points
                            y = pdf_height - (pos['top'] / 12700) - (pos['height'] / 12700)  # Flip Y-axis
                            width = pos['width'] / 12700  # Convert EMUs to points
                            height = pos['height'] / 12700  # Convert EMUs to points
                            
                            # Text frames
                            if shape.has_text_frame:
                                text = shape.text_frame.text
                                if text.strip():
                                    # Get text formatting
                                    formatting = self._get_text_formatting(shape)
                                    
                                    # Set font
                                    font_name = formatting['font_name']
                                    font_size = formatting['font_size']
                                    
                                    # Try to use registered Chinese fonts
                                    try:
                                        if "ChineseFont" in pdfmetrics.getRegisteredFontNames():
                                            font_name = "ChineseFont"
                                        elif "SimSun" in pdfmetrics.getRegisteredFontNames():
                                            font_name = "SimSun"
                                        elif "Microsoft YaHei" in pdfmetrics.getRegisteredFontNames():
                                            font_name = "Microsoft YaHei"
                                    except:
                                        font_name = "Helvetica"
                                    
                                    # Set font and color
                                    c.setFont(font_name, font_size)
                                    c.setFillColorRGB(*[c/255.0 for c in formatting['color']])
                                    
                                    # Process text line by line
                                    lines = text.split('\n')
                                    line_height = font_size * 1.2  # Line height is 1.2 times font size
                                    
                                    for i, line in enumerate(lines):
                                        if line.strip():
                                            # Calculate Y position for this line
                                            line_y = y + height - (i + 1) * line_height
                                            
                                            # Ensure text is properly encoded
                                            try:
                                                safe_line = line.encode('utf-8', errors='replace').decode('utf-8')
                                                c.drawString(x, line_y, safe_line)
                                            except Exception:
                                                safe_line = str(line).encode('utf-8', errors='replace').decode('utf-8')
                                                c.drawString(x, line_y, safe_line)
                            
                            # Tables
                            elif shape.has_table:
                                table = shape.table
                                table_data = []
                                
                                # Extract table data
                                for row_idx, row in enumerate(table.rows):
                                    row_data = []
                                    for cell_idx, cell in enumerate(row.cells):
                                        try:
                                            cell_text = cell.text_frame.text
                                            safe_text = cell_text.encode('utf-8', errors='replace').decode('utf-8')
                                            row_data.append(safe_text)
                                        except Exception:
                                            row_data.append("")
                                    table_data.append(row_data)
                                
                                # Create a ReportLab table
                                if table_data:
                                    # Calculate column widths
                                    cols = len(table_data[0])
                                    col_widths = [width / cols] * cols
                                    
                                    # Create table
                                    rl_table = Table(table_data, colWidths=col_widths)
                                    rl_table.setStyle(TableStyle([
                                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                        ('FONTNAME', (0, 0), (-1, 0), 'ChineseFont-Bold'),
                                        ('FONTSIZE', (0, 0), (-1, 0), 12),
                                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                        ('FONTNAME', (0, 1), (-1, -1), 'ChineseFont'),
                                        ('FONTSIZE', (0, 1), (-1, -1), 10),
                                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                    ]))
                                    
                                    # Draw the table
                                    table_width, table_height = rl_table.wrapOn(c, width, height)
                                    rl_table.drawOn(c, x, y)
                            
                            # Images
                            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                try:
                                    # Get image data
                                    image_bytes = shape.image.blob
                                    
                                    # Create a temporary image file
                                    img_stream = io.BytesIO(image_bytes)
                                    img_reader = ImageReader(img_stream)
                                    
                                    # Draw the image at the same position and size
                                    c.drawImage(img_reader, x, y, width=width, height=height)
                                except Exception as e:
                                    print(f"Error processing image: {str(e)}")
                                    continue
                            
                            # Other shapes (rectangles, circles, etc.)
                            else:
                                try:
                                    # Get fill color
                                    fill_color = colors.white
                                    if hasattr(shape, 'fill') and hasattr(shape.fill, 'fore_color'):
                                        if shape.fill.fore_color.type == MSO_THEME_COLOR_INDEX.RGB:
                                            fill_color = (
                                                shape.fill.fore_color.rgb.red / 255.0,
                                                shape.fill.fore_color.rgb.green / 255.0,
                                                shape.fill.fore_color.rgb.blue / 255.0
                                            )
                                    
                                    # Get line color
                                    line_color = colors.black
                                    line_width = 1
                                    if hasattr(shape, 'line') and hasattr(shape.line, 'color'):
                                        if shape.line.color.type == MSO_THEME_COLOR_INDEX.RGB:
                                            line_color = (
                                                shape.line.color.rgb.red / 255.0,
                                                shape.line.color.rgb.green / 255.0,
                                                shape.line.color.rgb.blue / 255.0
                                            )
                                        if hasattr(shape.line, 'width'):
                                            line_width = shape.line.width.pt
                                    
                                    # Set fill and line colors
                                    c.setFillColorRGB(*fill_color)
                                    c.setStrokeColorRGB(*line_color)
                                    c.setLineWidth(line_width)
                                    
                                    # Draw the shape based on its type
                                    if shape.shape_type == MSO_SHAPE_TYPE.RECTANGLE:
                                        c.rect(x, y, width, height, fill=1, stroke=1)
                                    elif shape.shape_type == MSO_SHAPE_TYPE.OVAL:
                                        c.ellipse(x, y, width, height, fill=1, stroke=1)
                                    elif shape.shape_type == MSO_SHAPE_TYPE.LINE:
                                        # For lines, we need start and end points
                                        if hasattr(shape, 'x1') and hasattr(shape, 'y1') and hasattr(shape, 'x2') and hasattr(shape, 'y2'):
                                            x1 = shape.x1 / 12700
                                            y1 = pdf_height - (shape.y1 / 12700)
                                            x2 = shape.x2 / 12700
                                            y2 = pdf_height - (shape.y2 / 12700)
                                            c.line(x1, y1, x2, y2)
                                        else:
                                            # Default to drawing a line from bottom-left to top-right of the shape bounds
                                            c.line(x, y, x + width, y + height)
                                    else:
                                        # For other shape types, draw a rectangle as fallback
                                        c.rect(x, y, width, height, fill=1, stroke=1)
                                except Exception as e:
                                    print(f"Error processing shape: {str(e)}")
                                    continue
                        except Exception as e:
                            print(f"Error processing shape: {str(e)}")
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
                "message": "PowerPoint presentation converted to PDF successfully with improved layout preservation".encode('utf-8', errors='replace').decode('utf-8'),
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}".encode('utf-8', errors='replace').decode('utf-8')}