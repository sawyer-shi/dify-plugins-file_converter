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
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader as RLImage
    from reportlab.platypus import Table as RLTable, TableStyle, Paragraph as RLParagraph, Spacer, SimpleDocTemplate, Image as RLImage2, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
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
    
    def _iter_block_items(self, parent):
        """
        按照文档中的实际顺序生成段落和表格对象
        这是关键函数，保证了图文混排的顺序不会被打乱
        
        Args:
            parent: Document对象或其他包含块级元素的对象
            
        Yields:
            Paragraph或Table对象，按照文档中的实际顺序
        """
        if hasattr(parent, 'element'):
            parent_elm = parent.element.body
        else:
            parent_elm = parent
            
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                # 段落元素
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                # 表格元素
                yield Table(child, parent)
    
    def _get_paragraph_alignment(self, paragraph):
        """
        获取段落的对齐方式
        
        Args:
            paragraph: python-docx的Paragraph对象
            
        Returns:
            reportlab的对齐常量
        """
        alignment = paragraph.alignment
        if alignment is None:
            return TA_LEFT
        elif alignment == 1:  # CENTER
            return TA_CENTER
        elif alignment == 2:  # RIGHT
            return TA_RIGHT
        elif alignment == 3:  # JUSTIFY
            return TA_JUSTIFY
        else:
            return TA_LEFT
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """
        使用python-docx和reportlab进行Word到PDF的转换
        关键改进：按照文档的实际顺序处理内容，保持图文混排
        """
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
            
            # Create PDF document with more appropriate margins
            pdf_doc = SimpleDocTemplate(
                output_path,
                pagesize=A4,
                rightMargin=50,
                leftMargin=50,
                topMargin=50,
                bottomMargin=30
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
            
            # Create custom styles for Chinese text with various sizes
            try:
                # 正文样式
                normal_style = ParagraphStyle(
                    'CustomNormal',
                    parent=styles['Normal'],
                    fontName=normal_font,
                    fontSize=11,
                    leading=16,
                    spaceAfter=8,
                    wordWrap='CJK',
                    alignment=TA_LEFT
                )
                
                # 标题样式
                heading1_style = ParagraphStyle(
                    'CustomHeading1',
                    parent=styles['Heading1'],
                    fontName=bold_font,
                    fontSize=18,
                    leading=22,
                    spaceAfter=12,
                    spaceBefore=12,
                    wordWrap='CJK',
                    alignment=TA_LEFT
                )
                
                heading2_style = ParagraphStyle(
                    'CustomHeading2',
                    parent=styles['Heading2'],
                    fontName=bold_font,
                    fontSize=16,
                    leading=20,
                    spaceAfter=10,
                    spaceBefore=10,
                    wordWrap='CJK',
                    alignment=TA_LEFT
                )
                
                heading3_style = ParagraphStyle(
                    'CustomHeading3',
                    parent=styles['Heading3'],
                    fontName=bold_font,
                    fontSize=14,
                    leading=18,
                    spaceAfter=8,
                    spaceBefore=8,
                    wordWrap='CJK',
                    alignment=TA_LEFT
                )
                
                # 居中样式
                center_style = ParagraphStyle(
                    'CustomCenter',
                    parent=normal_style,
                    alignment=TA_CENTER
                )
                
                # 右对齐样式
                right_style = ParagraphStyle(
                    'CustomRight',
                    parent=normal_style,
                    alignment=TA_RIGHT
                )
                
            except Exception:
                # Fallback to default styles if custom styles fail
                normal_style = styles['Normal']
                heading1_style = styles['Heading1']
                heading2_style = styles['Heading2']
                heading3_style = styles['Heading3']
                center_style = styles['Normal']
                right_style = styles['Normal']
            
            # Build PDF content - 关键：按文档顺序处理
            story = []
            
            # 收集所有图片的引用ID和内容
            image_parts = {}
            try:
                for rel in doc.part.rels.values():
                    if "image" in rel.target_ref:
                        try:
                            # 使用relationship ID作为键
                            image_parts[rel.rId] = rel.target_part.blob
                        except Exception:
                            continue
            except Exception:
                pass
            
            # 按顺序处理文档中的所有块级元素（段落和表格）
            for block in self._iter_block_items(doc):
                if isinstance(block, Paragraph):
                    # 处理段落
                    text = block.text.strip()
                    has_image = False
                    
                    # 先检查段落中是否有图片（即使没有文字也要检查！）
                    try:
                        for run in block.runs:
                            # 检查run中是否包含图片
                            if hasattr(run, '_element'):
                                for drawing in run._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'):
                                    # 尝试提取图片
                                    try:
                                        blip = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                                        if blip is not None:
                                            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                            if embed and embed in image_parts:
                                                has_image = True
                                                # 找到对应的图片
                                                image_data = image_parts[embed]
                                                image_stream = io.BytesIO(image_data)
                                                
                                                try:
                                                    # 创建图片对象
                                                    img = RLImage(image_stream)
                                                    img_width, img_height = img.getSize()
                                                    
                                                    # 计算合适的尺寸
                                                    max_width = 7 * inch
                                                    max_height = 9 * inch
                                                    
                                                    if img_width > max_width or img_height > max_height:
                                                        ratio = min(max_width / img_width, max_height / img_height)
                                                        img_width *= ratio
                                                        img_height *= ratio
                                                    
                                                    # 添加图片
                                                    # 重新创建BytesIO对象，因为前面读取过了
                                                    image_stream = io.BytesIO(image_data)
                                                    rl_img = RLImage2(image_stream, width=img_width, height=img_height)
                                                    story.append(Spacer(1, 6))
                                                    story.append(rl_img)
                                                    story.append(Spacer(1, 6))
                                                except Exception as img_err:
                                                    # 图片处理失败，记录但继续
                                                    print(f"Warning: Failed to process image: {img_err}")
                                                    continue
                                    except Exception as draw_err:
                                        print(f"Warning: Failed to extract image from drawing: {draw_err}")
                                        continue
                    except Exception as run_err:
                        print(f"Warning: Failed to process runs for images: {run_err}")
                        pass
                    
                    # 如果段落有文字，添加文字
                    if text:
                        try:
                            # 确定样式
                            style_name = block.style.name if block.style else 'Normal'
                            
                            if style_name.startswith('Heading 1'):
                                para_style = heading1_style
                            elif style_name.startswith('Heading 2'):
                                para_style = heading2_style
                            elif style_name.startswith('Heading 3'):
                                para_style = heading3_style
                            else:
                                # 根据对齐方式选择样式
                                alignment = self._get_paragraph_alignment(block)
                                if alignment == TA_CENTER:
                                    para_style = center_style
                                elif alignment == TA_RIGHT:
                                    para_style = right_style
                                else:
                                    para_style = normal_style
                            
                            # 创建段落
                            story.append(RLParagraph(text, para_style))
                                
                        except Exception as e:
                            # Fallback for text that can't be processed
                            print(f"Warning: Failed to process paragraph text: {e}")
                            try:
                                story.append(RLParagraph(text, normal_style))
                            except:
                                pass
                    elif not has_image:
                        # 既没有文字也没有图片的空段落，作为间距
                        story.append(Spacer(1, 6))
                
                elif isinstance(block, Table):
                    # 处理表格
                    try:
                        # Convert table data to list of lists
                        table_data = []
                        for row in block.rows:
                            row_data = []
                            for cell in row.cells:
                                cell_text = cell.text.strip()
                                # Handle empty cells
                                if not cell_text:
                                    cell_text = " "
                                row_data.append(cell_text)
                            table_data.append(row_data)
                        
                        if table_data:
                            try:
                                # 创建表格，让reportlab自动计算列宽
                                pdf_table = RLTable(table_data, repeatRows=1)
                                
                                # Add table style with better formatting
                                table_style = TableStyle([
                                    # 表头样式
                                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
                                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                    ('FONTNAME', (0, 0), (-1, 0), bold_font),
                                    ('FONTSIZE', (0, 0), (-1, 0), 11),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                                    ('TOPPADDING', (0, 0), (-1, 0), 10),
                                    # 数据行样式
                                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                    ('FONTNAME', (0, 1), (-1, -1), normal_font),
                                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                                    ('TOPPADDING', (0, 1), (-1, -1), 6),
                                    ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 8),
                                    ('RIGHTPADDING', (0, 0), (-1, -1), 8),
                                    # 边框
                                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                                    ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#4472C4')),
                                ])
                                
                                pdf_table.setStyle(table_style)
                                story.append(Spacer(1, 8))
                                story.append(pdf_table)
                                story.append(Spacer(1, 12))
                            except Exception as e:
                                # Fallback for table styling - 使用简单样式
                                print(f"Warning: Failed to apply table style, using simple style: {e}")
                                try:
                                    pdf_table = RLTable(table_data)
                                    story.append(Spacer(1, 8))
                                    story.append(pdf_table)
                                    story.append(Spacer(1, 12))
                                except Exception as e2:
                                    print(f"Warning: Failed to create table: {e2}")
                        else:
                            print("Warning: Empty table data, skipping table")
                    except Exception as e:
                        # Skip tables that can't be processed
                        print(f"Warning: Failed to process table: {e}")
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
                    "message": "Word document converted to PDF successfully using pure Python libraries (python-docx + reportlab) with improved layout preservation",
                    "output_files": output_files
                }
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts"}
                    
        except Exception as e:
            return {"success": False, "message": f"Error converting with python-docx and reportlab: {str(e)}"}
