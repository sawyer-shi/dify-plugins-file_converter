import os
import tempfile
from typing import Any, Dict, List, Optional, Generator
from collections.abc import Generator
from bs4 import BeautifulSoup

# WeasyPrint相关导入
try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File


class HtmlToPdfTool(Tool):
    """
    HTML to PDF conversion tool using WeasyPrint
    """
    
    def get_runtime_parameters(self) -> List[Dict[str, Any]]:
        return [
            {
                "name": "input_file",
                "type": "file",
                "required": True,
                "description": "要转换的HTML文件"
            },
            {
                "name": "output_file",
                "type": "string",
                "required": False,
                "description": "输出PDF文件名（可选，默认使用原文件名.pdf）"
            },
            {
                "name": "page_size",
                "type": "string",
                "required": False,
                "description": "页面大小（默认：A4）",
                "default": "A4"
            },
            {
                "name": "margin_top",
                "type": "string",
                "required": False,
                "description": "上边距（默认：1cm）",
                "default": "1cm"
            },
            {
                "name": "margin_bottom",
                "type": "string",
                "required": False,
                "description": "下边距（默认：1cm）",
                "default": "1cm"
            },
            {
                "name": "margin_left",
                "type": "string",
                "required": False,
                "description": "左边距（默认：1cm）",
                "default": "1cm"
            },
            {
                "name": "margin_right",
                "type": "string",
                "required": False,
                "description": "右边距（默认：1cm）",
                "default": "1cm"
            },
            {
                "name": "orientation",
                "type": "string",
                "required": False,
                "description": "页面方向（portrait/landscape，默认：portrait）",
                "default": "portrait"
            },
            {
                "name": "encoding",
                "type": "string",
                "required": False,
                "description": "HTML文件编码（默认utf-8）",
                "default": "utf-8"
            }
        ]

    def _invoke(self, tool_parameters: Dict[str, Any]) -> Generator[ToolInvokeMessage]:
        """
        Convert HTML to PDF using weasyprint
        """
        # 检查依赖
        if not WEASYPRINT_AVAILABLE:
            yield self.create_text_message("Error: weasyprint库未安装，请安装weasyprint>=57.0")
            return
        
        # 获取参数
        input_file = tool_parameters.get("input_file")
        output_file = tool_parameters.get("output_file")
        page_size = tool_parameters.get("page_size", "A4")
        margin = tool_parameters.get("margin", "2cm")
        encoding = tool_parameters.get("encoding", "utf-8")
        
        # 验证输入文件
        if not input_file:
            yield self.create_text_message("Error: 请提供有效的HTML文件")
            return
        
        try:
            # 读取HTML内容
            if hasattr(input_file, 'blob'):
                # 如果是File对象，使用blob属性
                html_content = input_file.blob
            elif hasattr(input_file, 'read'):
                # 如果是文件对象，使用read方法
                html_content = input_file.read()
            else:
                # 如果是字符串路径，直接读取文件
                with open(input_file, 'rb') as f:
                    html_content = f.read()
            
            # 如果是二进制内容，尝试解码
            if isinstance(html_content, bytes):
                try:
                    html_content = html_content.decode(encoding)
                except UnicodeDecodeError:
                    html_content = html_content.decode('utf-8', errors='replace')
            
            # 使用BeautifulSoup解析HTML
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 生成输出文件名
            if not output_file:
                if hasattr(input_file, 'filename'):
                    original_name = input_file.filename
                elif hasattr(input_file, 'name'):
                    original_name = input_file.name
                else:
                    original_name = 'html_document'
                base_name = os.path.splitext(original_name)[0]
                output_file = f"{base_name}.pdf"
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 设置输出文件完整路径
                output_path = os.path.join(temp_dir, output_file)
                
                # 添加CSS样式
                css_styles = self._css_styles()
                
                # 创建完整的HTML文档
                full_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="{encoding}">
                    <title>{soup.title.string if soup.title else 'Document'}</title>
                    <style>
                        {css_styles}
                    </style>
                </head>
                <body>
                    {soup.decode_contents()}
                </body>
                </html>
                """
                
                # 转换HTML到PDF
                HTML(string=full_html).write_pdf(
                    output_path,
                    stylesheets=[CSS(string=css_styles)]
                )
                
                # 获取文件大小
                file_size = os.path.getsize(output_path)
                file_size_mb = round(file_size / (1024 * 1024), 2)
                
                # 提取文本内容用于预览
                text_content = soup.get_text()
                preview_text = text_content[:500] + "..." if len(text_content) > 500 else text_content
                
                # 创建JSON响应
                json_response = {
                    "success": True,
                    "conversion_type": "html_2_pdf",
                    "input_file": {
                        "filename": getattr(input_file, 'filename', getattr(input_file, 'name', 'html_document')),
                        "size": len(html_content) if isinstance(html_content, bytes) else len(html_content.encode())
                    },
                    "output_file": {
                        "filename": output_file,
                        "size": file_size_mb,
                        "path": output_path
                    },
                    "preview_text": preview_text,
                    "message": f"HTML文件已成功转换为PDF格式\n输出文件: {output_file}\n文件大小: {file_size_mb} MB\n页面大小: {page_size}, 边距: {margin}"
                }
                
                # 发送文本消息
                yield self.create_text_message(f"HTML文件已成功转换为PDF格式: {output_file}")
                
                # 发送JSON消息
                yield self.create_json_message(json_response)
                
                # 发送文件
                with open(output_path, 'rb') as f:
                    yield self.create_blob_message(
                        blob=f.read(), 
                        meta={
                            "filename": output_file,
                            "mime_type": "application/pdf"
                        }
                    )
                
        except Exception as e:
            yield self.create_text_message(f"转换过程中发生错误: {str(e)}")
    
    def _css_styles(self):
        """创建CSS样式"""
        return """
        @page {
            size: A4 portrait;
            margin: 2cm;
        }
        
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 0;
        }
        
        h1, h2, h3, h4, h5, h6 {
            color: #222;
            margin-top: 1em;
            margin-bottom: 0.5em;
            font-weight: bold;
        }
        
        h1 { font-size: 2em; }
        h2 { font-size: 1.5em; }
        h3 { font-size: 1.17em; }
        h4 { font-size: 1em; }
        h5 { font-size: 0.83em; }
        h6 { font-size: 0.67em; }
        
        p {
            margin-bottom: 1em;
        }
        
        table {
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 1em;
        }
        
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        
        img {
            max-width: 100%;
            height: auto;
        }
        
        a {
            color: #0066cc;
            text-decoration: underline;
        }
        
        ul, ol {
            margin-bottom: 1em;
            padding-left: 2em;
        }
        
        li {
            margin-bottom: 0.5em;
        }
        """