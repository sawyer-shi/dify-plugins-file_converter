import os
import tempfile
import time
from collections.abc import Generator
from typing import Any, Dict, Optional
import json

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

try:
    import openpyxl
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    OPENPYXL_REPORTLAB_AVAILABLE = True
except ImportError:
    # Fallback for environments without openpyxl and reportlab
    OPENPYXL_REPORTLAB_AVAILABLE = False

class ExcelToPdfTool(Tool):
    """Tool for converting Excel documents to PDF format."""
    
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
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        try:
            # Get parameters
            file = tool_parameters.get("input_file")
            
            if not file:
                yield self.create_text_message("Error: Missing required parameter 'input_file'")
                return
                
            # Get file info
            file_info = self.get_file_info(file)
            if not file_info:
                yield self.create_text_message("Error: Invalid file")
                return
                
            # Validate input file format
            if not self._validate_input_file(file_info["extension"]):
                yield self.create_text_message("Error: Invalid file format. Only .xlsx files are supported (not .xls)")
                return
                
            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save the uploaded file to temp directory
                input_path = os.path.join(temp_dir, file_info["filename"])
                with open(input_path, "wb") as f:
                    f.write(file.blob)
                
                # Update file_info with the actual path
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
                        "conversion_type": "excel_2_pdf",
                        "input_file": file_info,
                        "output_files": output_files,
                        "message": result["message"]
                    }
                    
                    # Send text message
                    yield self.create_text_message(f"Excel converted to PDF successfully: {result['message']}")
                    
                    # Send JSON message
                    yield self.create_json_message(json_response)
                    
                    # Send output files
                    for output_file_info in result["output_files"]:
                        try:
                            # Use the pre-read content
                            if "content" in output_file_info:
                                yield self.create_blob_message(
                                    blob=output_file_info["content"], 
                                    meta={
                                        "filename": output_file_info["filename"],
                                        "mime_type": "application/pdf"
                                    }
                                )
                            else:
                                yield self.create_text_message(f"Error: No content available for file {output_file_info.get('filename', 'unknown')}")
                        except Exception as e:
                            yield self.create_text_message(f"Error sending file: {str(e)}")
                else:
                    # Send error message
                    yield self.create_text_message(f"Conversion failed: {result['message']}")
                    
        except Exception as e:
            yield self.create_text_message(f"Error during conversion: {str(e)}")
    
    def _validate_input_file(self, file_extension: str) -> bool:
        """Validate if the input file format is supported for Excel to PDF conversion."""
        return file_extension.lower() in [".xlsx"]
    
    def _register_chinese_fonts(self):
        """Register Chinese fonts for PDF generation."""
        try:
            # Get the directory of the current script
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # Get the parent directory (tools) and then the fonts directory
            fonts_dir = os.path.join(os.path.dirname(current_dir), "fonts")
            
            # Common Chinese font paths in Windows and project fonts
            font_paths = [
                # Project Chinese font (highest priority)
                (os.path.join(fonts_dir, "chinese_font.ttc"), "ChineseFont"),
                # SimSun (宋体)
                ("C:/Windows/Fonts/simsun.ttc", "SimSun"),
                ("C:/Windows/Fonts/simsun.ttf", "SimSun"),
                # SimHei (黑体)
                ("C:/Windows/Fonts/simhei.ttf", "SimHei"),
                # SimHei Bold - Note: Windows typically doesn't have a separate SimHei-Bold file
                # We'll use the same font file for both normal and bold styles
                ("C:/Windows/Fonts/simhei.ttf", "SimHei-Bold"),
                # Microsoft YaHei (微软雅黑)
                ("C:/Windows/Fonts/msyh.ttc", "MicrosoftYaHei"),
                ("C:/Windows/Fonts/msyh.ttf", "MicrosoftYaHei"),
                # Microsoft YaHei Bold
                ("C:/Windows/Fonts/msyhbd.ttc", "MicrosoftYaHei-Bold"),
                ("C:/Windows/Fonts/msyhbd.ttf", "MicrosoftYaHei-Bold"),
                # KaiTi (楷体)
                ("C:/Windows/Fonts/simkai.ttf", "KaiTi"),
                # FangSong (仿宋)
                ("C:/Windows/Fonts/simfang.ttf", "FangSong"),
                # SimSun Bold
                ("C:/Windows/Fonts/simsunb.ttf", "SimSun-Bold"),
            ]
            
            # Track if any Chinese fonts were successfully registered
            chinese_fonts_registered = False
            
            # Register fonts if they exist
            for font_path, font_name in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        chinese_fonts_registered = True
                    except Exception as e:
                        # Skip font registration if it fails
                        continue
            
            # If no Chinese fonts were registered, we'll use reportlab's built-in fonts
            # No need to register Helvetica as it's a built-in font
            return chinese_fonts_registered
                
        except Exception as e:
            # If font registration fails, we'll use reportlab's built-in fonts
            return False
    
    def _process_conversion(self, input_path: str, temp_dir: str) -> Dict[str, Any]:
        """Process the Excel to PDF conversion using openpyxl and reportlab."""
        output_files = []
        
        try:
            if not OPENPYXL_REPORTLAB_AVAILABLE:
                return {"success": False, "message": "openpyxl and reportlab libraries are not available for Excel conversion"}
            
            # Register Chinese fonts
            chinese_fonts_available = self._register_chinese_fonts()
            
            # Determine which fonts to use based on registration success
            if chinese_fonts_available:
                # Try to use Chinese fonts in order of preference
                try:
                    # Check if ChineseFont (project font) is available
                    pdfmetrics.getFont("ChineseFont")
                    normal_font = 'ChineseFont'
                    bold_font = 'ChineseFont'  # Use the same font for bold as well
                except:
                    try:
                        # Check if SimSun is available
                        pdfmetrics.getFont("SimSun")
                        normal_font = 'SimSun'
                        bold_font = 'SimSun-Bold'
                    except:
                        try:
                            # Check if Microsoft YaHei is available
                            pdfmetrics.getFont("MicrosoftYaHei")
                            normal_font = 'MicrosoftYaHei'
                            bold_font = 'MicrosoftYaHei-Bold'
                        except:
                            try:
                                # Check if SimHei is available
                                pdfmetrics.getFont("SimHei")
                                normal_font = 'SimHei'
                                bold_font = 'SimHei-Bold'
                            except:
                                # Fallback to built-in fonts
                                normal_font = 'Helvetica'
                                bold_font = 'Helvetica-Bold'
            else:
                # Use reportlab's built-in fonts
                normal_font = 'Helvetica'
                bold_font = 'Helvetica-Bold'
            
            # Generate output file path
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_path = os.path.join(temp_dir, f"{base_name}.pdf")
            
            # Load Excel workbook
            workbook = openpyxl.load_workbook(input_path)
            
            # Create PDF document
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            elements = []
            
            # Get styles for Chinese text
            styles = getSampleStyleSheet()
            chinese_style = ParagraphStyle(
                'ChineseStyle',
                parent=styles['Normal'],
                fontName=normal_font,
                fontSize=10,
                leading=14
            )
            
            heading_style = ParagraphStyle(
                'ChineseHeading',
                parent=styles['Heading1'],
                fontName=bold_font,
                fontSize=14,
                leading=18,
                alignment=TA_CENTER
            )
            
            # Process each sheet in the workbook
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Extract data from sheet
                data = []
                for row in sheet.iter_rows(values_only=True):
                    # Convert None to empty string for display
                    data_row = []
                    for cell in row:
                        if cell is not None:
                            # Convert to string and ensure proper encoding for Chinese text
                            cell_str = str(cell)
                            # Ensure the string is properly encoded for Chinese characters
                            if isinstance(cell_str, str):
                                # No need to encode/decode, just ensure it's a proper string
                                data_row.append(cell_str)
                            else:
                                # Convert to string if it's not already
                                data_row.append(str(cell_str))
                        else:
                            data_row.append("")
                    data.append(data_row)
                
                # Create table from data
                if data:
                    # Calculate column widths based on content
                    col_count = len(data[0]) if data else 0
                    col_widths = []
                    
                    for col_idx in range(col_count):
                        max_length = 0
                        for row in data:
                            if col_idx < len(row) and row[col_idx]:
                                cell_value = str(row[col_idx])
                                max_length = max(max_length, len(cell_value))
                        # Set column width with minimum and maximum limits
                        col_width = min(max(max_length * 0.1, 0.5), 2.0) * inch
                        col_widths.append(col_width)
                    
                    # Create table with style
                    table = Table(data, colWidths=col_widths)
                    
                    # Add table style with Chinese font support
                    style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), bold_font),
                        ('FONTSIZE', (0, 0), (-1, 0), 12),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('FONTNAME', (0, 1), (-1, -1), normal_font),
                        ('FONTSIZE', (0, 1), (-1, -1), 10),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ])
                    
                    table.setStyle(style)
                    
                    # Add sheet name as heading with Chinese font support
                    heading = Paragraph(f"<b>{sheet_name}</b>", heading_style)
                    elements.append(heading)
                    elements.append(Spacer(1, 0.2 * inch))
                    
                    # Add table to elements
                    elements.append(table)
                    elements.append(Spacer(1, 0.5 * inch))
            
            # Build PDF document
            doc.build(elements)
            
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
            else:
                return {"success": False, "message": "Failed to read converted file after multiple attempts"}
            
            return {
                "success": True, 
                "message": "Excel spreadsheet converted to PDF successfully with Chinese font support",
                "output_files": output_files
            }
                
        except Exception as e:
            return {"success": False, "message": f"Conversion error: {str(e)}"}