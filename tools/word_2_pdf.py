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
                yield self.create_text_message("Error: Invalid file format. Only .docx files are supported (not .doc)")
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
        if not file_info["extension"].lower().endswith('.docx'):
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
    
    def _parse_numbering_format(self, doc):
        """
        解析Word文档的编号格式定义
        从numbering.xml中提取所有编号格式信息
        
        Args:
            doc: Document对象
            
        Returns:
            字典：{(numId, level): {'format': 'decimal'/'chineseCounting'/etc, 'prefix': '', 'suffix': ''}}
        """
        numbering_formats = {}
        
        try:
            numbering_part = doc.part.numbering_part
            if numbering_part is None:
                return numbering_formats
            
            # 获取numbering元素
            numbering_element = numbering_part.element
            
            # 解析abstractNum定义（抽象编号定义）
            abstract_nums = {}
            for abstractNum in numbering_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum'):
                abstractNumId = abstractNum.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
                
                # 解析每个级别的定义
                for lvl in abstractNum.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl'):
                    ilvl = lvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl')
                    
                    # 获取编号格式
                    numFmt = lvl.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt')
                    fmt = numFmt.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if numFmt is not None else 'decimal'
                    
                    # 获取编号文本格式（包含前缀、后缀）
                    lvlText = lvl.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText')
                    text_format = lvlText.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if lvlText is not None else '%1'
                    
                    # 获取起始值
                    start = lvl.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start')
                    start_val = int(start.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')) if start is not None else 1
                    
                    if abstractNumId not in abstract_nums:
                        abstract_nums[abstractNumId] = {}
                    
                    abstract_nums[abstractNumId][ilvl] = {
                        'format': fmt,
                        'text_format': text_format,
                        'start': start_val
                    }
            
            # 解析num定义（实例化的编号定义）
            for num in numbering_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num'):
                numId = num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
                
                # 获取关联的abstractNum
                abstractNumId = num.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId')
                if abstractNumId is not None:
                    abstractNumIdVal = abstractNumId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                    
                    # 复制abstractNum的定义到num
                    if abstractNumIdVal in abstract_nums:
                        for ilvl, fmt_info in abstract_nums[abstractNumIdVal].items():
                            numbering_formats[(int(numId), int(ilvl))] = fmt_info
            
            return numbering_formats
            
        except Exception as e:
            print(f"Warning: Failed to parse numbering formats: {e}")
            return numbering_formats
    
    def _format_number(self, count, format_type):
        """
        根据格式类型格式化编号
        
        Args:
            count: 编号计数值
            format_type: 编号格式类型（decimal、chineseCounting、lowerLetter等）
            
        Returns:
            格式化后的编号字符串
        """
        if format_type == 'decimal':
            return str(count)
        elif format_type == 'chineseCounting':
            # 中文数字编号
            chinese_nums = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
                          "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
                          "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十"]
            if count < len(chinese_nums):
                return chinese_nums[count]
            else:
                return str(count)
        elif format_type == 'chineseCountingThousand':
            # 大写中文数字
            return self._convert_to_chinese_number(count)
        elif format_type == 'lowerLetter':
            # 小写字母 a, b, c...
            if count <= 26:
                return chr(ord('a') + count - 1)
            else:
                return str(count)
        elif format_type == 'upperLetter':
            # 大写字母 A, B, C...
            if count <= 26:
                return chr(ord('A') + count - 1)
            else:
                return str(count)
        elif format_type == 'lowerRoman':
            # 小写罗马数字
            return self._to_roman(count).lower()
        elif format_type == 'upperRoman':
            # 大写罗马数字
            return self._to_roman(count)
        elif format_type == 'bullet':
            # 项目符号
            return '•'
        else:
            # 默认使用阿拉伯数字
            return str(count)
    
    def _convert_to_chinese_number(self, num):
        """转换为中文大写数字"""
        if num == 0:
            return "零"
        
        digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
        units = ["", "十", "百", "千"]
        
        if num < 10:
            return digits[num]
        elif num < 20:
            return "十" + (digits[num - 10] if num > 10 else "")
        elif num < 100:
            tens = num // 10
            ones = num % 10
            return digits[tens] + "十" + (digits[ones] if ones > 0 else "")
        else:
            # 简化处理，只处理到99
            return str(num)
    
    def _to_roman(self, num):
        """转换为罗马数字"""
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syms[i]
                num -= val[i]
            i += 1
        return roman_num
    
    def _extract_numbering_text(self, paragraph):
        """
        尝试从段落的XML中提取编号文本的替代方法
        检查段落是否以常见的编号模式开头
        
        Args:
            paragraph: python-docx的Paragraph对象
            
        Returns:
            提取的编号文本，如果没有则返回空字符串
        """
        try:
            # 获取段落的完整文本
            full_text = paragraph.text
            
            # 检查是否以常见的编号模式开头
            import re
            
            # 匹配中文数字编号：一、二、三、
            chinese_pattern = r'^([一二三四五六七八九十百千]+)[、．.]'
            match = re.match(chinese_pattern, full_text)
            if match:
                return match.group(0) + " "
            
            # 匹配阿拉伯数字编号：1. 2. 3.
            arabic_pattern = r'^(\d+)[、．.]'
            match = re.match(arabic_pattern, full_text)
            if match:
                return match.group(0) + " "
            
            # 匹配带括号的编号：(1) (2) (3) 或 （1）（2）（3）
            paren_pattern = r'^[（(]\d+[）)]'
            match = re.match(paren_pattern, full_text)
            if match:
                return match.group(0) + " "
            
            return ""
            
        except Exception:
            return ""
    
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
            
            # 解析文档的编号格式定义
            print("Parsing numbering formats from document...")
            numbering_formats = self._parse_numbering_format(doc)
            print(f"Found {len(numbering_formats)} numbering format definitions")
            
            # 创建编号计数器字典，用于追踪不同级别的编号
            # key: (numId, level), value: 当前计数
            numbering_counters = {}
            
            # 按顺序处理文档中的所有块级元素（段落和表格）
            for block in self._iter_block_items(doc):
                if isinstance(block, Paragraph):
                    # 处理段落 - 更精细地提取文本，保留所有字符
                    # 使用paragraph.text会自动过滤一些内容，我们需要更精确的提取
                    text_parts = []
                    for run in block.runs:
                        if run.text:
                            text_parts.append(run.text)
                    
                    # 合并所有run的文本
                    text = ''.join(text_parts).strip()
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
                            # 检查段落是否有自动编号
                            numbering_text = ""
                            try:
                                if block._element.pPr is not None:
                                    numPr = block._element.pPr.numPr
                                    if numPr is not None:
                                        # 段落有编号属性
                                        numId_elem = numPr.numId
                                        ilvl_elem = numPr.ilvl
                                        
                                        if numId_elem is not None:
                                            numId = numId_elem.val
                                            
                                            # 关键检查：numId为0表示段落没有实际编号
                                            # 这是Word中取消编号的方式
                                            if numId == 0:
                                                # 这个段落没有编号，跳过
                                                print(f"Skipping numbering: numId=0 (no numbering)")
                                            else:
                                                # numId > 0，这是一个真正有编号的段落
                                                level = ilvl_elem.val if ilvl_elem is not None else 0
                                                
                                                # 更新计数器
                                                counter_key = (numId, level)
                                                
                                                # 获取起始值
                                                start_val = 1
                                                if counter_key in numbering_formats:
                                                    start_val = numbering_formats[counter_key].get('start', 1)
                                                
                                                if counter_key not in numbering_counters:
                                                    numbering_counters[counter_key] = start_val - 1
                                                numbering_counters[counter_key] += 1
                                                
                                                # 生成编号文本
                                                count = numbering_counters[counter_key]
                                                
                                                # 使用从文档中解析的格式
                                                if counter_key in numbering_formats:
                                                    fmt_info = numbering_formats[counter_key]
                                                    format_type = fmt_info['format']
                                                    text_format = fmt_info['text_format']
                                                    
                                                    # 格式化编号
                                                    formatted_num = self._format_number(count, format_type)
                                                    
                                                    # 应用文本格式（替换%1, %2等占位符）
                                                    # text_format例如："%1、" 或 "第%1章" 或 "(%1)"
                                                    numbering_text = text_format.replace(f'%{level + 1}', formatted_num)
                                                    
                                                    # 如果text_format中没有占位符，直接使用格式化的数字
                                                    if '%' not in text_format:
                                                        numbering_text = formatted_num + " "
                                                    
                                                    print(f"Generated numbering: numId={numId}, level={level}, count={count}, format={format_type}, text={numbering_text}")
                                                else:
                                                    # 如果没有找到格式定义，使用默认格式
                                                    print(f"Warning: No format found for numId={numId}, level={level}, using default")
                                                    numbering_text = f"{count}. "
                            except Exception as num_err:
                                # 编号提取失败，继续处理
                                print(f"Warning: Failed to generate numbering: {num_err}")
                                pass
                            
                            # 如果提取到编号，添加到文本前面
                            if numbering_text:
                                text = numbering_text + text
                            
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
                        # 获取表格的列数
                        if not block.rows:
                            continue
                        
                        num_cols = len(block.rows[0].cells)
                        
                        # 创建用于表格单元格的样式（左对齐，自动换行）
                        cell_style = ParagraphStyle(
                            'TableCell',
                            parent=normal_style,
                            fontSize=9,
                            leading=12,
                            wordWrap='CJK',
                            alignment=TA_LEFT,  # 左对齐
                            leftIndent=0,
                            rightIndent=0
                        )
                        
                        # 转换表格数据为Paragraph对象列表，支持自动换行
                        table_data = []
                        for row_idx, row in enumerate(block.rows):
                            row_data = []
                            for cell in row.cells:
                                cell_text = cell.text.strip()
                                # Handle empty cells
                                if not cell_text:
                                    cell_text = " "
                                
                                # 使用Paragraph包装文本，支持自动换行和中文显示
                                try:
                                    para = RLParagraph(cell_text, cell_style)
                                    row_data.append(para)
                                except Exception as e:
                                    # 如果Paragraph创建失败，使用纯文本
                                    row_data.append(cell_text)
                            
                            table_data.append(row_data)
                        
                        if table_data:
                            try:
                                # 计算可用宽度（A4纸宽度 - 左右边距）
                                page_width = A4[0]
                                available_width = page_width - 100  # 减去左右边距
                                
                                # 根据列数平均分配列宽
                                col_width = available_width / num_cols
                                col_widths = [col_width] * num_cols
                                
                                # 创建表格，指定列宽
                                pdf_table = RLTable(
                                    table_data, 
                                    colWidths=col_widths,
                                    repeatRows=1
                                )
                                
                                # Add table style with better formatting
                                table_style = TableStyle([
                                    # 表头样式
                                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
                                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                    ('ALIGN', (0, 0), (-1, 0), 'LEFT'),  # 表头也左对齐
                                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 顶部对齐
                                    ('FONTNAME', (0, 0), (-1, 0), bold_font),
                                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                                    ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                                    ('TOPPADDING', (0, 0), (-1, 0), 8),
                                    # 数据行样式
                                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                    ('ALIGN', (0, 1), (-1, -1), 'LEFT'),  # 左对齐
                                    ('VALIGN', (0, 1), (-1, -1), 'TOP'),  # 顶部对齐
                                    ('FONTNAME', (0, 1), (-1, -1), normal_font),
                                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                                    ('TOPPADDING', (0, 1), (-1, -1), 5),
                                    ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
                                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                                    # 边框
                                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                                    ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor('#4472C4')),
                                    # 自动换行
                                    ('WORDWRAP', (0, 0), (-1, -1), True),
                                ])
                                
                                pdf_table.setStyle(table_style)
                                story.append(Spacer(1, 8))
                                story.append(pdf_table)
                                story.append(Spacer(1, 12))
                            except Exception as e:
                                # Fallback for table styling - 使用简单样式
                                print(f"Warning: Failed to apply table style, using simple style: {e}")
                                try:
                                    # 简化版：使用纯文本，平均列宽
                                    simple_data = []
                                    for row in block.rows:
                                        row_data = []
                                        for cell in row.cells:
                                            cell_text = cell.text.strip() if cell.text.strip() else " "
                                            row_data.append(cell_text)
                                        simple_data.append(row_data)
                                    
                                    pdf_table = RLTable(simple_data, colWidths=col_widths)
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
