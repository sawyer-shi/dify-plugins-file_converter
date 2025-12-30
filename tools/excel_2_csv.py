import os
import tempfile
import zipfile
from typing import Any, Dict, Generator

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# 导入依赖库，包含错误处理
try:
    import pandas as pd
    import openpyxl
    import xlrd
    DEPENDENCIES_AVAILABLE = True
except ImportError:
    DEPENDENCIES_AVAILABLE = False

class ExcelToCsvTool(Tool):
    """
    Excel to CSV Converter Tool.
    Converts Excel (.xlsx) files to CSV format, with support for multiple worksheets.
    Each worksheet will be converted to a separate CSV file.
    """

    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        if not DEPENDENCIES_AVAILABLE:
            yield self.create_text_message("Error: Required libraries (pandas, openpyxl, xlrd) are missing.")
            return

        input_file = tool_parameters.get("input_file")
        if not input_file:
            yield self.create_text_message("Error: Input file is required.")
            return

        # 验证文件格式
        if not input_file.extension or input_file.extension.lower() not in ['.xlsx', '.xls']:
            yield self.create_text_message(f"Error: Only .xlsx or .xls files are supported. Provided file extension: {input_file.extension or 'None'}")
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
        
        # 验证文件是否看起来像Excel文件
        try:
            with tempfile.NamedTemporaryFile(suffix=input_file.extension) as temp_file:
                temp_file.write(input_file.blob)
                temp_file.flush()
                
                # 使用pandas尝试读取文件来验证
                pd.ExcelFile(temp_file.name)
        except Exception:
            yield self.create_text_message("Warning: File may not be a valid Excel file.")
        
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 准备输入文件
                input_path = os.path.join(temp_dir, input_file.filename)
                with open(input_path, "wb") as f:
                    f.write(input_file.blob)
                
                # 执行转换
                converter = ExcelCsvConverter(input_path, temp_dir)
                result = converter.convert()

                if not result["success"]:
                    yield self.create_text_message(f"Conversion Failed: {result['message']}")
                    return

                # 返回转换结果信息
                yield self.create_text_message(f"Successfully converted Excel to CSV files:\n{result['message']}")
                
                # 返回所有生成的CSV文件
                for csv_file in result["files"]:
                    with open(csv_file["path"], 'rb') as f:
                        csv_content = f.read()
                    
                    yield self.create_blob_message(
                        blob=csv_content,
                        meta={
                            "filename": csv_file["name"],
                            "mime_type": "text/csv"
                        }
                    )

        except Exception as e:
            yield self.create_text_message(f"System Error: {str(e)}")

class ExcelCsvConverter:
    """
    内部转换器类，负责具体的Excel到CSV转换
    """
    def __init__(self, input_path: str, output_dir: str):
        self.input_path = input_path
        self.output_dir = output_dir

    def convert(self) -> Dict[str, Any]:
        try:
            # 读取Excel文件
            excel_file = pd.ExcelFile(self.input_path)
            sheet_names = excel_file.sheet_names
            
            if not sheet_names:
                return {"success": False, "message": "Excel file contains no worksheets"}
            
            converted_files = []
            total_rows = 0
            
            # 获取基础文件名（不含扩展名）
            base_filename = os.path.splitext(os.path.basename(self.input_path))[0]
            
            for sheet_name in sheet_names:
                try:
                    # 读取工作表数据
                    df = pd.read_excel(self.input_path, sheet_name=sheet_name)
                    
                    # 将自动生成的 "Unnamed: X" 列名替换为空字符串
                    df.columns = ['' if str(col).startswith('Unnamed: ') else col for col in df.columns]
                    
                    # 清理工作表名称，用作文件名
                    safe_sheet_name = self._sanitize_filename(sheet_name)
                    
                    # 生成CSV文件名
                    if safe_sheet_name.lower() == base_filename.lower():
                        # 如果工作表名与文件名相同，直接使用基础文件名
                        csv_filename = f"{base_filename}.csv"
                    else:
                        # 否则组合基础文件名和工作表名
                        csv_filename = f"{base_filename}_{safe_sheet_name}.csv"
                    
                    csv_path = os.path.join(self.output_dir, csv_filename)
                    
                    # 保存为CSV文件
                    df.to_csv(csv_path, index=False, encoding='utf-8')
                    
                    # 获取行数和列数
                    rows, cols = df.shape
                    total_rows += rows
                    
                    converted_files.append({
                        "name": csv_filename,
                        "path": csv_path,
                        "sheet_name": sheet_name,
                        "rows": rows,
                        "cols": cols
                    })
                    
                except Exception as e:
                    # 如果某个工作表转换失败，记录错误但继续处理其他工作表
                    print(f"Warning: Failed to convert sheet '{sheet_name}': {str(e)}")
                    continue
            
            if not converted_files:
                return {"success": False, "message": "Failed to convert any worksheets"}
            
            # 生成结果消息
            message = f"Converted {len(converted_files)} worksheet(s) to CSV files:\n"
            for file_info in converted_files:
                message += f"- {file_info['name']}: {file_info['rows']} rows, {file_info['cols']} columns (from sheet '{file_info['sheet_name']}')\n"
            message += f"Total: {total_rows} rows across all worksheets"
            
            return {
                "success": True,
                "message": message,
                "files": converted_files
            }

        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "message": str(e)}
    
    def _sanitize_filename(self, name: str) -> str:
        """
        清理工作表名称，确保可以用作文件名
        """
        if not name:
            return "sheet"
        
        # 移除或替换不允许的文件名字符
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # 移除多余的空格和点
        name = name.strip(' .')
        
        # 如果清理后为空，使用默认名称
        if not name:
            name = "sheet"
        
        return name