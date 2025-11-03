import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# Try to import python-docx and reportlab components for Word to PDF conversion
try:
    from docx import Document
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader as RLImage
    from reportlab.platypus import Table, TableStyle, Paragraph, Spacer, SimpleDocTemplate
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    import io
    DOCX_REPORTLAB_AVAILABLE = True
except ImportError:
    DOCX_REPORTLAB_AVAILABLE = False

# Try to import reportlab font components for Chinese font support
try:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.fonts import addMapping
    REPORTLAB_FONT_AVAILABLE = True
except ImportError:
    REPORTLAB_FONT_AVAILABLE = False

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
    
    def _register_chinese_fonts(self):
        """Register Chinese fonts for reportlab to use."""
        if not REPORTLAB_FONT_AVAILABLE:
            return False
            
        try:
            registered_fonts = []
            
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
                ('Microsoft YaHei', 'C:/Windows/Fonts/msyhbd.ttf'),  # Bold variant
                # KaiTi (楷体)
                ('KaiTi', 'C:/Windows/Fonts/kaiti.ttf'),
                # FangSong (仿宋)
                ('FangSong', 'C:/Windows/Fonts/simfang.ttf'),
            ]
            
            for font_name, font_path in font_paths:
                try:
                    if os.path.exists(font_path):
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        registered_fonts.append(font_name)
                except Exception as e:
                    # Continue trying other fonts if one fails
                    continue
            
            # Register bold variants if available
            bold_variants = [
                ('SimSun-Bold', 'C:/Windows/Fonts/simsunb.ttf'),
                ('SimHei-Bold', 'C:/Windows/Fonts/simheib.ttf'),
                ('Microsoft YaHei-Bold', 'C:/Windows/Fonts/msyhbd.ttf'),
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
                    'SimSun': 'Helvetica',
                    'SimHei': 'Helvetica',
                    'SimSun-Bold': 'Helvetica-Bold',
                    'SimHei-Bold': 'Helvetica-Bold',
                    'Microsoft YaHei': 'Helvetica',
                    'Microsoft YaHei-Bold': 'Helvetica-Bold',
                    'KaiTi': 'Helvetica',
                    'FangSong': 'Helvetica',
                }
                
                for chinese_font, fallback_font in font_mapping.items():
                    try:
                        # Create an alias for the fallback font
                        addMapping(chinese_font, 0, 0, fallback_font)
                        addMapping(chinese_font, 1, 0, fallback_font)
                        addMapping(chinese_font, 0, 1, fallback_font)
                        addMapping(chinese_font, 1, 1, fallback_font)
                    except Exception:
                        # If even fallback fails, just continue
                        continue
                        
            return len(registered_fonts) > 0
                        
        except Exception as e:
            # If font registration fails completely, we'll rely on default fonts
            return False
    
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
        """Validate if the input file is a valid Word document."""
        # Check file extension
        if not file_info["extension"].lower().endswith(('.docx', '.doc')):
            return False
            
        # Check if file is readable by python-docx
        if DOCX_REPORTLAB_AVAILABLE and "path" in file_info:
            try:
                doc = Document(file_info["path"])
                # Try to access the document to ensure it's valid
                _ = doc.paragraphs[0].text if doc.paragraphs else ""
                return True
            except Exception:
                return False
        
        # If python-docx is not available or path not available, just check file extension
        return True
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """Process the Word to PDF conversion using python-docx and reportlab."""
        output_files = []
        
        # Generate output file path
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, f"{base_name}.pdf")
        
        # Check if required libraries are available
        if not DOCX_REPORTLAB_AVAILABLE:
            return {"success": False, "message": "Required libraries (python-docx, reportlab) are not available. Please install them using: pip install python-docx reportlab"}
        
        try:
            # Register Chinese fonts for reportlab
            chinese_fonts_registered = self._register_chinese_fonts()
            
            # Load the Word document
            doc = Document(input_path)
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(
                output_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # Get styles
            styles = getSampleStyleSheet()
            
            # Determine which fonts to use based on registration success
            if chinese_fonts_registered:
                # Try to use Chinese fonts in order of preference
                try:
                    # Check if ChineseFont is available (project font)
                    pdfmetrics.getFont("ChineseFont")
                    normal_font = 'ChineseFont'
                    bold_font = 'ChineseFont'  # Use same font for bold
                except:
                    try:
                        # Check if SimSun is available
                        pdfmetrics.getFont("SimSun")
                        normal_font = 'SimSun'
                        bold_font = 'SimSun-Bold'
                    except:
                        try:
                            # Check if Microsoft YaHei is available
                            pdfmetrics.getFont("Microsoft YaHei")
                            normal_font = 'Microsoft YaHei'
                            bold_font = 'Microsoft YaHei-Bold'
                        except:
                            # Fallback to any available Chinese font
                            normal_font = 'SimHei'
                            bold_font = 'SimHei'
            else:
                # Use reportlab's built-in fonts
                normal_font = 'Helvetica'
                bold_font = 'Helvetica-Bold'
            
            # Create custom styles for Chinese text
            try:
                normal_style = ParagraphStyle(
                    'CustomNormal',
                    parent=styles['Normal'],
                    fontName=normal_font,
                    fontSize=10,
                    leading=14,
                    spaceAfter=6,
                    wordWrap='CJK'
                )
                
                heading_style = ParagraphStyle(
                    'CustomHeading',
                    parent=styles['Heading1'],
                    fontName=bold_font,
                    fontSize=14,
                    leading=18,
                    spaceAfter=12,
                    wordWrap='CJK'
                )
            except Exception:
                # Fallback to default styles if custom styles fail
                normal_style = styles['Normal']
                heading_style = styles['Heading1']
            
            # Build PDF content
            story = []
            
            # Process paragraphs
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    try:
                        # Try to determine if this is a heading
                        if para.style.name.startswith('Heading'):
                            story.append(Paragraph(text, heading_style))
                        else:
                            story.append(Paragraph(text, normal_style))
                    except Exception as e:
                        # Fallback for text that can't be processed
                        story.append(Paragraph(text, normal_style))
            
            # Process tables
            for table in doc.tables:
                try:
                    # Convert table data to list of lists
                    table_data = []
                    for row in table.rows:
                        row_data = []
                        for cell in row.cells:
                            cell_text = cell.text.strip()
                            # Handle empty cells
                            if not cell_text:
                                cell_text = " "
                            row_data.append(cell_text)
                        table_data.append(row_data)
                    
                    if table_data:
                        # Create table with appropriate style
                        try:
                            pdf_table = Table(table_data)
                            
                            # Add table style
                            table_style = TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), bold_font),
                                ('FONTSIZE', (0, 0), (-1, 0), 10),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                ('FONTNAME', (0, 1), (-1, -1), normal_font),
                                ('FONTSIZE', (0, 1), (-1, -1), 9),
                            ])
                            
                            pdf_table.setStyle(table_style)
                            story.append(pdf_table)
                            story.append(Spacer(1, 12))
                        except Exception as e:
                            # Fallback for table styling
                            pdf_table = Table(table_data)
                            story.append(pdf_table)
                            story.append(Spacer(1, 12))
                except Exception as e:
                    # Skip tables that can't be processed
                    continue
            
            # Process images
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    try:
                        image_data = rel.target_part.blob
                        image_stream = io.BytesIO(image_data)
                        img = RLImage(image_stream)
                        
                        # Calculate image size to fit page
                        img_width, img_height = img.getSize()
                        max_width = 6 * inch  # Max width is 6 inches
                        max_height = 8 * inch  # Max height is 8 inches
                        
                        # Scale image if necessary
                        if img_width > max_width or img_height > max_height:
                            ratio = min(max_width / img_width, max_height / img_height)
                            img_width *= ratio
                            img_height *= ratio
                        
                        # Create a reportlab Image object instead of using ImageReader directly
                        from reportlab.platypus import Image as RLImagePlatypus
                        rl_img = RLImagePlatypus(image_stream, width=img_width, height=img_height)
                        
                        # Add image to story
                        story.append(Spacer(1, 6))
                        story.append(rl_img)
                        story.append(Spacer(1, 6))
                    except Exception as e:
                        # Skip images that can't be processed
                        continue
            
            # Build PDF
            pdf_doc.build(story)
            
            # Wait for file to be fully written
            time.sleep(2)
            
            # Check if file exists and has content
            if not os.path.exists(output_path):
                return {"success": False, "message": "Output PDF file was not created"}
                
            if os.path.getsize(output_path) == 0:
                return {"success": False, "message": "Output PDF file is empty"}
            
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
                    "message": "Word document converted to PDF successfully using python-docx and reportlab with Chinese font support",
                    "output_files": output_files
                }
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts"}
                    
        except Exception as e:
            return {"success": False, "message": f"Error converting with python-docx and reportlab: {str(e)}"}