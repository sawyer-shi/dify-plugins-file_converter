import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional, List, Tuple
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# Try to import PyMuPDF for PDF text extraction
try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

# Try to import pdfplumber as alternative for PDF text extraction
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

class PdfToTextTool(Tool):
    """Tool for converting PDF files to text format with enhanced table support."""
    
    def get_file_info(self, file: File) -> dict:
        file_info = {
            "filename": file.filename,
            "extension": file.extension,
            "mime_type": file.mime_type,
            "size": file.size,
            "url": file.url
        }
        if hasattr(file, 'path'):
            file_info["path"] = file.path
        return file_info
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            file = tool_parameters.get("input_file")
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            file_info = self.get_file_info(file)
            
            if not self._validate_input_file(file_info):
                yield self.create_text_message("Error: Invalid file format. Only PDF files (.pdf) are supported")
                return
                
            with tempfile.TemporaryDirectory() as temp_dir:
                # 保存文件的逻辑
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, 'wb') as f:
                    f.write(file.blob)
                
                file_info["path"] = input_path
                
                # 执行转换
                result = self._process_conversion(input_path, temp_dir)
                
                if result["success"]:
                    # 构建返回结果... (保持原有逻辑不变)
                    output_files = []
                    for output_file_info in result["output_files"]:
                        output_files.append({
                            "filename": output_file_info["filename"],
                            "size": len(output_file_info["content"]),
                            "path": output_file_info["path"]
                        })
                    
                    json_response = {
                        "success": True,
                        "conversion_type": "pdf_2_text",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    yield self.create_text_message(f"PDF file converted to text successfully: {result['message']}")
                    yield self.create_json_message(json_response)
                    
                    for file_info in result["output_files"]:
                        if "content" in file_info:
                            yield self.create_blob_message(
                                blob=file_info["content"], 
                                meta={"filename": file_info["filename"], "mime_type": "text/plain"}
                            )
                else:
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _validate_input_file(self, file_info: dict) -> bool:
        if not file_info["extension"].lower().endswith('.pdf'):
            return False
        return True
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """Process the PDF conversion using the best available method for tables."""
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        output_path = os.path.join(temp_dir, f"{base_name}.txt")
        
        if not PYMUPDF_AVAILABLE and not PDFPLUMBER_AVAILABLE:
            return {"success": False, "message": "Required library (PyMuPDF or pdfplumber) is not available."}
        
        try:
            text_content = ""
            method_used = "Unknown"

            # 优先使用 PyMuPDF 的高级表格检测逻辑
            if PYMUPDF_AVAILABLE:
                try:
                    # 检查 fitz 版本是否支持 find_tables (v1.23.0+)
                    doc = fitz.open(input_path)
                    if hasattr(doc[0], "find_tables"): 
                        text_content = self._extract_with_pymupdf_tables(doc)
                        method_used = "PyMuPDF (Table Detection)"
                    else:
                        # 回退到普通提取
                        for page in doc:
                            text_content += page.get_text() + "\n\n"
                        method_used = "PyMuPDF (Standard)"
                    doc.close()
                except Exception as e:
                    print(f"PyMuPDF extraction failed, falling back: {e}")
                    if PDFPLUMBER_AVAILABLE:
                        text_content = self._extract_with_pdfplumber(input_path)
                        method_used = "pdfplumber"
                    else:
                        raise e
            elif PDFPLUMBER_AVAILABLE:
                # 如果只有 pdfplumber，也可以尝试表格提取逻辑（虽然比 fitz 慢）
                text_content = self._extract_with_pdfplumber(input_path)
                method_used = "pdfplumber"

            # 写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(text_content)
            
            # 读取文件用于返回
            with open(output_path, 'r', encoding='utf-8') as f:
                file_content = f.read()

            return {
                "success": True, 
                "message": f"Converted using {method_used}",
                "output_files": [{"path": output_path, "content": file_content.encode('utf-8'), "filename": f"{base_name}.txt"}]
            }
                    
        except Exception as e:
            return {"success": False, "message": f"Error converting PDF: {str(e)}"}

    def _extract_with_pymupdf_tables(self, doc) -> str:
        """
        使用 PyMuPDF 提取文本，并智能识别表格，将其转换为 'word2text' 风格的 Markdown 表格。
        核心逻辑：检测表格区域 -> 提取表格数据 -> 提取非表格区域文本 -> 按页面垂直位置重组。
        """
        full_text = []

        for page_num, page in enumerate(doc):
            if page_num > 0:
                full_text.append(f"\n\n--- Page {page_num + 1} ---\n\n")

            # 1. 查找表格
            tabs = page.find_tables()
            
            # 存储所有需要按顺序输出的内容块：{'y0': float, 'text': str, 'type': 'table'|'text'}
            content_blocks = []

            # 处理表格
            existing_tables_bboxes = [] # 用于后续剔除重复文本
            for tab in tabs:
                # 记录表格区域以便稍后过滤文本
                existing_tables_bboxes.append(fitz.Rect(tab.bbox))
                
                # 格式化表格
                table_lines = ["\n--- Table ---"]
                # tab.extract() 返回 [['col1', 'col2'], ...]
                table_data = tab.extract() 
                if not table_data: 
                    continue
                    
                for row in table_data:
                    # 清洗单元格内容：去除换行，替换None为""
                    clean_row = []
                    for cell in row:
                        if cell is None:
                            clean_row.append("")
                        else:
                            # 将单元格内的换行符替换为空格，保持单行
                            clean_row.append(str(cell).replace('\n', ' ').strip())
                    
                    # 使用管道符连接
                    table_lines.append(" | ".join(clean_row))
                
                table_lines.append("--- End of Table ---\n")
                formatted_table = "\n".join(table_lines)
                
                content_blocks.append({
                    'y0': tab.bbox[1], # 表格顶部y坐标
                    'text': formatted_table,
                    'type': 'table'
                })

            # 2. 提取文本块
            # 使用 "blocks" 模式获取坐标：(x0, y0, x1, y1, "text", block_no, block_type)
            text_blocks_raw = page.get_text("blocks")
            
            for block in text_blocks_raw:
                # block[4] 是文本内容，block[6] 是类型(0=text, 1=image)
                if block[6] != 0: 
                    continue
                    
                b_text = block[4].strip()
                if not b_text:
                    continue
                    
                b_rect = fitz.Rect(block[0], block[1], block[2], block[3])
                
                # 检查该文本块是否在某个表格内
                is_inside_table = False
                for t_rect in existing_tables_bboxes:
                    # 如果文本块和表格区域交集超过文本块面积的 70%，认为它是表格的一部分，跳过
                    intersect = b_rect & t_rect # 交集
                    if intersect.get_area() / b_rect.get_area() > 0.7:
                        is_inside_table = True
                        break
                
                if not is_inside_table:
                    content_blocks.append({
                        'y0': block[1],
                        'text': b_text,
                        'type': 'text'
                    })

            # 3.按垂直位置(y0)排序，重组页面内容
            # 主要按 y0 排序，次要按 x0 (对于多栏布局可能需要更复杂的逻辑，但通常 block 顺序已经涵盖)
            content_blocks.sort(key=lambda x: x['y0'])

            for block in content_blocks:
                full_text.append(block['text'])
        
        return "\n".join(full_text)

    def _extract_with_pdfplumber(self, input_path: str) -> str:
        """
        Fallback: Use pdfplumber. Modified to try to extract tables explicitly if possible.
        """
        text_content = []
        try:
            with pdfplumber.open(input_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    if page_num > 0:
                        text_content.append(f"\n\n--- Page {page_num + 1} ---\n\n")
                    
                    # 尝试提取表格
                    extracted_tables = page.extract_tables()
                    
                    if extracted_tables:
                        # 如果有表格，这变得比较复杂，因为pdfplumber不容易混合文本和表格位置。
                        # 简单策略：先输出表格数据的类似格式，再输出剩余文本（可能会重复）
                        # 或者为了为了保持 word2text 格式，我们就只用 page.extract_text() 
                        # 但 pdfplumber 默认不加 |。
                        # 因此，我们对 pdfplumber 做一个简单的增强：优先展示表格。
                        
                        # 注意：在 pdfplumber 中混合流式文本比较慢且难。
                        # 这里我们做一个折中：如果有表格，先列出格式化表格，
                        # 然后使用 filter_outside_tables 提取剩余文本。
                        
                        # 1. 格式化表格
                        for table in extracted_tables:
                            text_content.append("--- Table ---")
                            for row in table:
                                clean_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                                text_content.append(" | ".join(clean_row))
                            text_content.append("--- End of Table ---\n")
                            
                        # 2. 尝试提取非表格文本 (需要找到表格的 bbox，稍微复杂，这里做简化处理)
                        # 简单追加全文（可能会有重复内容，但保证了信息不丢）
                        # 或者仅做标准提取：
                        text_content.append(page.extract_text() or "")
                    else:
                        # 无表格，直接提取
                        text_content.append(page.extract_text() or "")
                        
        except Exception as e:
            raise Exception(f"Error extracting text with pdfplumber: {str(e)}")
        
        return "\n".join(text_content)