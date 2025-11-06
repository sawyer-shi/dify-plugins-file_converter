import os
import tempfile
from collections.abc import Generator
from typing import Any, Dict, Optional
import json
import time
import io

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    import fitz  # PyMuPDF
    from docx import Document
    from docx.shared import Inches
    PYPDF2_AVAILABLE = True
except ImportError:
    # Fallback for environments without required libraries
    PYPDF2_AVAILABLE = False

class PdfToWordTool(Tool):
    """Tool for converting PDF documents to Word format."""
    
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
            output_format = tool_parameters.get("output_format", "docx")
            
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
                # Save uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, 'wb') as f:
                    f.write(file.blob)
                
                # Update file info with the actual path
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
                        "conversion_type": "pdf_2_word",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"PDF converted to Word successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for file_path in result["output_files"]:
                        yield self.create_blob_message(
                            blob=open(file_path, 'rb').read(), 
                            meta={
                                "filename": os.path.basename(file_path),
                                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            }
                        )
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _format_table(self, word_table, table_data):
        """
        Format the Word table with proper styling and preserve headers.
        改进：更简洁的格式，避免过度装饰
        """
        from docx.shared import Pt, Inches
        from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import nsmap
        from docx.oxml import OxmlElement
        
        # Fill the table with data and apply formatting
        for row_idx, row in enumerate(table_data):
            for col_idx, cell_data in enumerate(row):
                cell = word_table.cell(row_idx, col_idx)
                
                # 清空默认内容
                cell.text = ""
                
                # 添加内容
                if cell_data is not None and str(cell_data).strip():
                    p = cell.paragraphs[0]
                    run = p.add_run(str(cell_data))
                    
                    # 设置字体大小
                    run.font.size = Pt(10)
                    
                    # Format header row (first row) differently
                    if row_idx == 0:
                        # Make header text bold
                        run.font.bold = True
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Set header row background to light blue (更接近PDF样式)
                        try:
                            shading_elm = OxmlElement('w:shd')
                            shading_elm.set('{%s}val' % nsmap['w'], 'clear')
                            shading_elm.set('{%s}color' % nsmap['w'], 'auto')
                            shading_elm.set('{%s}fill' % nsmap['w'], 'D0E4F7')  # 淡蓝色
                            cell._element.get_or_add_tcPr().append(shading_elm)
                        except Exception as e:
                            print(f"Warning: Failed to apply cell shading: {e}")
                    else:
                        # Align content to left for data rows
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # 设置单元格垂直居中
                    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        # Set table alignment to left (更自然)
        word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        
        # 设置表格自动调整
        word_table.autofit = True
        
        # 设置表格样式为网格线
        try:
            word_table.style = 'Table Grid'
        except:
            pass
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for PDF to Word conversion."""
        return file_extension.lower() == ".pdf"
    
    def _get_elements_with_position(self, page, pdf_document):
        """
        获取页面所有元素及其位置，按照从上到下的顺序排列
        返回: [(y_position, element_type, element_data), ...]
        """
        elements = []
        
        # 1. 获取文本块及其位置
        try:
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            
            for block in blocks:
                if "lines" in block:
                    # 文本块
                    bbox = block.get("bbox", (0, 0, 0, 0))
                    y_position = bbox[1]  # top y coordinate
                    x_position = bbox[0]  # left x coordinate (用于检测缩进)
                    
                    # 提取文本内容和格式
                    block_text = ""
                    font_size = 12  # 默认字体大小
                    is_bold = False
                    color = None  # RGB颜色
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span.get("text", "")
                            # 获取字体信息
                            if "size" in span:
                                font_size = span["size"]
                            if "flags" in span:
                                # flags & 16 表示粗体
                                is_bold = (span["flags"] & 16) != 0
                            # 获取颜色信息 (RGB格式)
                            if "color" in span:
                                color = span["color"]
                        block_text += line_text + "\n"
                    
                    if block_text.strip():
                        elements.append((
                            y_position,
                            "text",
                            {
                                "text": block_text.strip(),
                                "font_size": font_size,
                                "is_bold": is_bold,
                                "color": color,
                                "x_position": x_position,
                                "bbox": bbox
                            }
                        ))
        except Exception as e:
            print(f"Warning: Failed to extract text blocks: {e}")
        
        # 2. 获取表格及其位置
        try:
            tables = page.find_tables()
            if tables.tables:
                for table in tables.tables:
                    try:
                        bbox = table.bbox
                        y_position = bbox[1]  # top y coordinate
                        table_data = table.extract()
                        
                        if table_data and len(table_data) > 0:
                            elements.append((
                                y_position,
                                "table",
                                {
                                    "data": table_data,
                                    "bbox": bbox
                                }
                            ))
                    except Exception as e:
                        print(f"Warning: Failed to extract table: {e}")
                        continue
        except Exception as e:
            print(f"Warning: Failed to find tables: {e}")
        
        # 3. 获取图片及其位置
        try:
            image_list = page.get_images()
            for img_index, img in enumerate(image_list):
                try:
                    # 获取图片位置
                    xref = img[0]
                    # 获取图片在页面上的位置
                    img_rects = page.get_image_rects(xref)
                    
                    if img_rects:
                        bbox = img_rects[0]
                        y_position = bbox[1]  # top y coordinate
                        
                        # 获取图片数据
                        pix = fitz.Pixmap(pdf_document, xref)
                        
                        # Skip CMYK images
                        if pix.n - pix.alpha < 4:
                            img_data = pix.tobytes("png")
                            
                            elements.append((
                                y_position,
                                "image",
                                {
                                    "data": img_data,
                                    "bbox": bbox,
                                    "width": pix.width,
                                    "height": pix.height
                                }
                            ))
                        
                        pix = None
                except Exception as e:
                    print(f"Warning: Failed to extract image {img_index}: {e}")
                    continue
        except Exception as e:
            print(f"Warning: Failed to get images: {e}")
        
        # 按照y坐标排序（从上到下）
        elements.sort(key=lambda x: x[0])
        
        return elements
    
    def _process_conversion(self, file_info: Dict[str, Any], output_format: str, temp_dir: str) -> Dict[str, Any]:
        """
        Process the PDF to Word conversion using PyMuPDF library.
        关键改进：按照PDF的实际布局顺序（从上到下）处理内容，不添加额外内容
        """
        input_path = file_info["path"]
        output_files = []
        
        try:
            if not PYPDF2_AVAILABLE:
                return {"success": False, "message": "Required libraries (PyMuPDF, python-docx) are not available for PDF conversion"}
            
            # Default to docx if not specified
            if not output_format:
                output_format = "docx"
            elif output_format.lower() not in ["doc", "docx"]:
                output_format = "docx"
            
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.{output_format}")
            
            # Open the PDF file
            try:
                pdf_document = fitz.open(input_path)
            except Exception as e:
                return {"success": False, "message": f"Failed to open PDF file: {str(e)}"}
            
            # Create a new Word document
            doc = Document()
            
            # Process each page
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                
                # 获取页面所有元素并按位置排序
                elements = self._get_elements_with_position(page, pdf_document)
                
                # 按顺序处理每个元素
                for y_pos, element_type, element_data in elements:
                    if element_type == "text":
                        # 添加文本段落
                        from docx.shared import Pt, RGBColor, Inches
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        
                        text = element_data["text"]
                        font_size = element_data["font_size"]
                        is_bold = element_data["is_bold"]
                        color = element_data.get("color")
                        x_position = element_data.get("x_position", 0)
                        
                        # 判断是否为标题（根据字体大小）
                        if font_size >= 16:
                            # 大字体，作为一级标题
                            heading = doc.add_heading(text, level=1)
                            # 应用颜色到标题
                            if color is not None:
                                for run in heading.runs:
                                    # PDF颜色是整数，需要转换为RGB
                                    r = (color >> 16) & 0xFF
                                    g = (color >> 8) & 0xFF
                                    b = color & 0xFF
                                    run.font.color.rgb = RGBColor(r, g, b)
                        elif font_size >= 14:
                            # 中等字体，作为二级标题
                            heading = doc.add_heading(text, level=2)
                            # 应用颜色到标题
                            if color is not None:
                                for run in heading.runs:
                                    r = (color >> 16) & 0xFF
                                    g = (color >> 8) & 0xFF
                                    b = color & 0xFF
                                    run.font.color.rgb = RGBColor(r, g, b)
                        else:
                            # 普通文本
                            p = doc.add_paragraph(text)
                            
                            # 设置缩进（根据x坐标）
                            # 页面左边距通常是72点（1英寸），大于这个值说明有缩进
                            if x_position > 80:  # 有明显缩进
                                indent_inches = (x_position - 72) / 72.0  # 转换为英寸
                                p.paragraph_format.left_indent = Inches(min(indent_inches, 2.0))  # 限制最大缩进
                            
                            # 应用格式到run
                            for run in p.runs:
                                if is_bold:
                                    run.bold = True
                                # 应用颜色
                                if color is not None:
                                    r = (color >> 16) & 0xFF
                                    g = (color >> 8) & 0xFF
                                    b = color & 0xFF
                                    run.font.color.rgb = RGBColor(r, g, b)
                                # 设置字体大小
                                run.font.size = Pt(max(font_size, 9))  # 最小9pt
                    
                    elif element_type == "table":
                        # 添加表格
                        table_data = element_data["data"]
                        
                        if table_data and len(table_data) > 0:
                            cols = len(table_data[0])
                            word_table = doc.add_table(rows=len(table_data), cols=cols)
                            word_table.style = 'Table Grid'
                            
                            # 填充表格数据（不添加额外说明）
                            self._format_table(word_table, table_data)
                    
                    elif element_type == "image":
                        # 添加图片
                        img_data = element_data["data"]
                        img_stream = io.BytesIO(img_data)
                        
                        # 根据原始尺寸计算合适的显示宽度
                        width = element_data["width"]
                        height = element_data["height"]
                        
                        # 限制最大宽度为6英寸
                        max_width = 6.0
                        if width > height:
                            doc_width = min(max_width, width / 100.0)  # 简单的缩放
                        else:
                            doc_width = min(4.0, width / 100.0)
                        
                        doc.add_picture(img_stream, width=Inches(doc_width))
                
                # 在页面之间添加分页符（除了最后一页）
                if page_num < len(pdf_document) - 1:
                    doc.add_page_break()
            
            # Close the PDF document
            pdf_document.close()
            
            # Save the Word document
            try:
                doc.save(output_path)
            except Exception as e:
                return {"success": False, "message": f"Failed to save Word document: {str(e)}"}
            
            # Check if file exists and has content
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                output_files.append(output_path)
            else:
                return {"success": False, "message": "Output file was not created or is empty"}
            
            return {
                "success": True, 
                "message": f"PDF converted to Word ({output_format}) successfully - preserving original layout order",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}