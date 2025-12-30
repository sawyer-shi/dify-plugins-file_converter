import os
import tempfile
from typing import Any, Dict, Generator

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# 导入依赖库，包含错误处理
try:
    import pandas as pd
    import openpyxl
    DEPENDENCIES_AVAILABLE = True
except ImportError:
    DEPENDENCIES_AVAILABLE = False

class CsvToExcelTool(Tool):
    """
    CSV to Excel Converter Tool.
    Converts CSV files to Excel (.xlsx) format using pandas and openpyxl.
    """

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        if not DEPENDENCIES_AVAILABLE:
            yield self.create_text_message("Error: Required libraries (pandas, openpyxl) are missing.")
            return

        input_file = tool_parameters.get("input_file")
        if not input_file:
            yield self.create_text_message("Error: Input file is required.")
            return

        # 验证文件格式
        if not input_file.extension or input_file.extension.lower() not in ['.csv']:
            yield self.create_text_message(f"Error: Only .csv files are supported. Provided file extension: {input_file.extension or 'None'}")
            return
        
        # 验证文件内容不为空
        if not input_file.blob or len(input_file.blob) == 0:
            yield self.create_text_message("Error: Input file is empty.")
            return
        
        # 验证文件大小（限制为50MB）
        max_file_size = 50 * 1024 * 1024  # 50MB
        if len(input_file.blob) > max_file_size:
            yield self.create_text_message(f"Error: File size exceeds maximum limit of 50MB. Current size: {len(input_file.blob) / (1024*1024):.2f}MB")
            return
        
        # 验证文件内容是否看起来像CSV格式
        try:
            # 读取文件的前1KB进行基本检查
            file_header = input_file.blob[:1024].decode('utf-8', errors='ignore')
            # 检查是否包含常见的CSV分隔符
            common_separators = [',', ';', '\t', '|']
            has_csv_structure = any(sep in file_header for sep in common_separators)
            
            if not has_csv_structure:
                yield self.create_text_message("Warning: File may not be in CSV format. No common CSV separators (comma, semicolon, tab, pipe) found in the file header.")
        except Exception:
            # 如果无法解码，继续处理，让pandas处理编码问题
            pass

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 准备输入文件
                input_path = os.path.join(temp_dir, input_file.filename)
                with open(input_path, "wb") as f:
                    f.write(input_file.blob)
                
                # 准备输出文件路径
                output_filename = os.path.splitext(input_file.filename)[0] + ".xlsx"
                output_path = os.path.join(temp_dir, output_filename)

                # 执行转换
                converter = CsvExcelConverter(input_path, output_path)
                result = converter.convert()

                if not result["success"]:
                    yield self.create_text_message(f"Conversion Failed: {result['message']}")
                    return

                # 读取并返回结果
                with open(output_path, 'rb') as f:
                    excel_content = f.read()

                yield self.create_text_message(f"Successfully converted CSV to Excel: {output_filename}\n{result['message']}")
                
                yield self.create_blob_message(
                    blob=excel_content,
                    meta={
                        "filename": output_filename,
                        "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    }
                )

        except Exception as e:
            yield self.create_text_message(f"System Error: {str(e)}")

class CsvExcelConverter:
    """
    内部转换器类，负责具体的CSV到Excel转换
    """
    def __init__(self, input_path: str, output_path: str):
        self.input_path = input_path
        self.output_path = output_path

    def convert(self) -> Dict[str, Any]:
        try:
            # 首先检查文件是否为空
            if os.path.getsize(self.input_path) == 0:
                return {"success": False, "message": "CSV file is empty"}
            
            # 尝试不同的编码方式读取CSV文件
            encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'iso-8859-1']
            df = None
            used_encoding = None
            
            for encoding in encodings:
                try:
                    # 使用pandas读取CSV，但允许空文件
                    df = pd.read_csv(self.input_path, encoding=encoding)
                    used_encoding = encoding
                    break
                except UnicodeDecodeError:
                    continue
                except pd.errors.EmptyDataError:
                    # 处理空文件情况 - 直接返回错误，不创建Excel文件
                    return {"success": False, "message": "CSV file is empty"}
                except Exception as e:
                    # 如果不是编码问题，直接返回错误
                    return {"success": False, "message": f"Error reading CSV: {str(e)}"}
            
            if df is None:
                return {"success": False, "message": "Unable to read CSV file with any supported encoding"}
            
            # 检查DataFrame是否为空
            if df.empty:
                # 如果DataFrame为空，直接返回错误，不创建Excel文件
                return {"success": False, "message": "CSV file is empty"}
            
            # 获取原始CSV文件的行数和列数
            rows, cols = df.shape
            
            # 从输入文件名生成工作表名称
            base_filename = os.path.basename(self.input_path)
            sheet_name = os.path.splitext(base_filename)[0]
            
            # 确保工作表名称符合Excel规范（不超过31个字符，不包含特殊字符）
            sheet_name = self._sanitize_sheet_name(sheet_name)
            
            # 使用ExcelWriter写入Excel文件
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                
                # 获取工作簿和工作表对象以进行格式化
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                
                # 自动调整列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)  # 设置最大宽度为50
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            return {
                "success": True, 
                "message": f"Excel file created successfully with {rows} rows and {cols} columns using {used_encoding} encoding, sheet name: '{sheet_name}'"
            }

        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "message": str(e)}
    
    def _sanitize_sheet_name(self, name: str) -> str:
        """
        清理工作表名称，确保符合Excel规范
        Excel工作表名称限制：
        - 最多31个字符
        - 不能包含以下字符: \\ / ? * [ ] 
        - 不能为空
        """
        if not name:
            return "Sheet1"
        
        # 移除不允许的字符
        invalid_chars = ['\\', '/', '?', '*', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '')
        
        # 限制长度
        if len(name) > 31:
            name = name[:31]
        
        # 如果清理后为空，使用默认名称
        if not name.strip():
            name = "Sheet1"
        
        return name