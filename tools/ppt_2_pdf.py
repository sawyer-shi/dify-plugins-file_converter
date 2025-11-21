import os
import tempfile
import io
import time
from typing import Any, Dict, List, Optional, Union

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# 导入依赖库，包含错误处理
try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    # 【修复点】这里导入 MSO_COLOR_TYPE
    from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
    from pptx.dml.color import RGBColor
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Table, TableStyle, Paragraph, Frame, KeepInFrame
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    DEPENDENCIES_AVAILABLE = True
except ImportError:
    DEPENDENCIES_AVAILABLE = False

# 常量定义
EMU_PER_INCH = 914400
PT_PER_INCH = 72
EMU_TO_PT = PT_PER_INCH / EMU_PER_INCH

class PptToPdfTool(Tool):
    """
    Advanced PPT to PDF Converter (Pure Python).
    Features: Recursive Group Shapes, Text Wrapping, Exact Table Sizing.
    """

    def _invoke(self, tool_parameters: dict[str, Any]) -> Any:
        if not DEPENDENCIES_AVAILABLE:
            yield self.create_text_message("Error: Required libraries (python-pptx, reportlab, Pillow) are missing.")
            return

        input_file = tool_parameters.get("input_file")
        if not input_file:
            yield self.create_text_message("Error: Input file is required.")
            return

        if not input_file.extension or input_file.extension.lower() not in ['.pptx']:
            yield self.create_text_message("Error: Only .pptx files are supported (not old .ppt).")
            return

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # 准备文件
                input_path = os.path.join(temp_dir, input_file.filename)
                with open(input_path, "wb") as f:
                    f.write(input_file.blob)
                
                output_filename = os.path.splitext(input_file.filename)[0] + ".pdf"
                output_path = os.path.join(temp_dir, output_filename)

                # 执行核心转换
                converter = PptPdfEngine(input_path, output_path)
                result = converter.convert()

                if not result["success"]:
                    yield self.create_text_message(f"Conversion Failed: {result['message']}")
                    return

                # 返回结果
                with open(output_path, 'rb') as f:
                    pdf_content = f.read()

                yield self.create_text_message("PPT conversion successful.")
                yield self.create_blob_message(
                    blob=pdf_content,
                    meta={
                        "filename": output_filename,
                        "mime_type": "application/pdf"
                    }
                )

        except Exception as e:
            import traceback
            traceback.print_exc()
            yield self.create_text_message(f"System Error: {str(e)}")

class PptPdfEngine:
    """
    PPT转换引擎核心类
    """
    def __init__(self, input_path: str, output_path: str):
        self.input_path = input_path
        self.output_path = output_path
        self.font_name = "CustomChineseFont"
        self.font_bold_name = "CustomChineseFont" 
        self._register_fonts()

    def _register_fonts(self):
        """注册字体，路径策略与Excel插件保持一致"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # 假设字体在 tools 同级的 fonts 目录
            font_path = os.path.join(os.path.dirname(current_dir), "fonts", "chinese_font.ttc")
            
            if not os.path.exists(font_path):
                # 备用：当前目录 fonts/
                font_path = os.path.join(current_dir, "fonts", "chinese_font.ttc")

            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont(self.font_name, font_path))
                # 简单的粗体映射
                pdfmetrics.registerFont(TTFont(self.font_bold_name, font_path)) 
            else:
                print("Warning: Chinese font not found, falling back to Helvetica")
                self.font_name = "Helvetica"
                self.font_bold_name = "Helvetica-Bold"
        except Exception as e:
            print(f"Font registration error: {e}")
            self.font_name = "Helvetica"
            self.font_bold_name = "Helvetica-Bold"

    def convert(self) -> Dict[str, Any]:
        try:
            prs = Presentation(self.input_path)
            
            # 获取PPT尺寸并转换为PDF点数
            slide_width_pt = prs.slide_width * EMU_TO_PT
            slide_height_pt = prs.slide_height * EMU_TO_PT
            
            c = canvas.Canvas(self.output_path, pagesize=(slide_width_pt, slide_height_pt))
            
            for slide in prs.slides:
                # 绘制每张幻灯片
                self._process_slide(c, slide, slide_height_pt)
                c.showPage()
            
            c.save()
            return {"success": True, "message": "OK"}
        except Exception as e:
            import traceback
            traceback.print_exc()
            # 打印更详细的错误堆栈以便调试
            return {"success": False, "message": f"Convert Error: {str(e)}"}

    def _process_slide(self, c: canvas.Canvas, slide: Any, page_height: float):
        """处理单个幻灯片"""
        # 1. 绘制背景
        try:
            bg = slide.background
            # 只有当背景填充明确定义时才绘制
            if bg and bg.fill and bg.fill.type: 
                color = self._get_solid_fill_color(bg.fill)
                if color:
                    c.setFillColor(color)
                    c.rect(0, 0, c._pagesize[0], c._pagesize[1], fill=1, stroke=0)
        except:
            pass

        # 2. 递归绘制所有形状
        if slide.shapes:
            for shape in slide.shapes:
                self._render_shape_recursive(c, shape, 0, 0, page_height)

    def _render_shape_recursive(self, c: canvas.Canvas, shape: Any, x_offset: float, y_offset: float, page_height: float):
        """
        递归渲染形状
        """
        # 跳过不可见元素
        if hasattr(shape, 'visible') and not shape.visible:
            return

        # 获取绝对坐标 (EMU)
        try:
             # 如果是组合里的子元素，left/top 是相对组合左上角的
            current_x_emu = x_offset + shape.left
            current_y_emu = y_offset + shape.top
        except AttributeError:
            return

        # 1. 处理组合形状 (Group)
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in shape.shapes:
                self._render_shape_recursive(c, sub_shape, current_x_emu, current_y_emu, page_height)
            return

        # 转换坐标为 PDF Point
        x = current_x_emu * EMU_TO_PT
        w = shape.width * EMU_TO_PT
        h = shape.height * EMU_TO_PT
        y = page_height - (current_y_emu * EMU_TO_PT) - h 

        # 2. 处理文本框 (Text Box)
        if shape.has_text_frame and shape.text_frame.text.strip():
            self._draw_shape_background(c, shape, x, y, w, h)
            self._draw_smart_text_box(c, shape, x, y, w, h)
        
        # 3. 处理表格 (Table)
        elif shape.has_table:
            self._draw_exact_table(c, shape.table, x, y, w, h, page_height)

        # 4. 处理图片 (Picture)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                image_blob = shape.image.blob
                img_reader = ImageReader(io.BytesIO(image_blob))
                # mask='auto' 处理透明背景png
                c.drawImage(img_reader, x, y, width=w, height=h, mask='auto', preserveAspectRatio=True)
            except Exception as e:
                print(f"Image render error: {e}")
                
        # 5. 其他几何形状
        elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
             self._draw_shape_background(c, shape, x, y, w, h)

    def _draw_shape_background(self, c: canvas.Canvas, shape: Any, x, y, w, h):
        """绘制形状背景和边框"""
        fill_color = None
        line_color = None
        line_width = 0

        # 获取填充颜色
        if hasattr(shape, 'fill'):
            fill_color = self._get_solid_fill_color(shape.fill)

        # 获取边框颜色
        if hasattr(shape, 'line') and shape.line.fill: # 注意：line.fill 才是颜色源
             line_color = self._get_solid_fill_color(shape.line.fill)
             if hasattr(shape.line, 'width') and shape.line.width:
                 line_width = shape.line.width.pt

        if fill_color or (line_color and line_width > 0):
            if fill_color:
                c.setFillColor(fill_color)
            else:
                c.setFillColor(colors.Color(0,0,0,alpha=0)) # 透明填充

            if line_color and line_width > 0:
                c.setStrokeColor(line_color)
                c.setLineWidth(line_width)
            else:
                c.setStrokeColor(colors.Color(0,0,0,alpha=0)) # 无边框
            
            # 绘制矩形（简化处理几何形状）
            c.rect(x, y, w, h, fill=1 if fill_color else 0, stroke=1 if line_color else 0)

    def _draw_smart_text_box(self, c: canvas.Canvas, shape: Any, x, y, w, h):
        """
        使用 ReportLab Paragraph 实现自动换行的文本框
        """
        text_frame = shape.text_frame
        styles = getSampleStyleSheet()
        flowables = []
        
        for paragraph in text_frame.paragraphs:
            # 即使是空行也需要保留占位
            if not paragraph.text and not paragraph.runs:
                flowables.append(Paragraph("<br/>", styles["Normal"]))
                continue

            font_size = 12
            font_color = colors.black
            
            # 尝试从第一个 run 获取样式
            if paragraph.runs:
                run = paragraph.runs[0]
                if run.font.size:
                    font_size = run.font.size.pt
                
                # 【修复点】使用 MSO_COLOR_TYPE 判断
                if run.font.color and run.font.color.type == MSO_COLOR_TYPE.RGB:
                     font_color = self._rgb_to_color(run.font.color.rgb)
            
            # 样式定义
            style = ParagraphStyle(
                name=f'P_{id(paragraph)}',
                parent=styles['Normal'],
                fontName=self.font_name,
                fontSize=font_size,
                textColor=font_color,
                leading=font_size * 1.2,
                wordWrap='CJK' 
            )
            
            # 文本内容清洗
            raw_text = paragraph.text if paragraph.text else ""
            if not raw_text.strip():
                # 如果只有空格或空
                flowables.append(Paragraph("<br/>", style))
                continue
                
            text_content = raw_text.replace('\n', '<br/>')
            text_content = text_content.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            
            flowables.append(Paragraph(text_content, style))

        if not flowables:
            return

        # 计算 Frame 内补白
        padding = 2
        # 使用 Frame + KeepInFrame 控制边界
        frame = Frame(x, y, w, h, showBoundary=0, topPadding=padding, leftPadding=padding, rightPadding=padding, bottomPadding=padding)
        story = [KeepInFrame(w - 2*padding, h - 2*padding, flowables, mode='shrink')]
        frame.addFromList(story, c)

    def _draw_exact_table(self, c: canvas.Canvas, ppt_table: Any, x, y, w, h, page_height):
        """绘制表格"""
        if not ppt_table.rows:
            return

        data = []
        row_heights = []
        
        for row in ppt_table.rows:
            row_data = []
            row_heights.append(row.height * EMU_TO_PT)
            for cell in row.cells:
                text = ""
                if cell.text_frame and cell.text_frame.text:
                    text = cell.text_frame.text.strip()
                row_data.append(text)
            data.append(row_data)

        col_widths = [col.width * EMU_TO_PT for col in ppt_table.columns]

        processed_data = []
        base_style = ParagraphStyle(
            name='TableBase', 
            fontName=self.font_name, 
            fontSize=10, 
            leading=11,
            textColor=colors.black,
            wordWrap='CJK'
        )

        for row in data:
            new_row = []
            for cell_text in row:
                safe_text = cell_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                new_row.append(Paragraph(safe_text, base_style))
            processed_data.append(new_row)

        rl_table = Table(processed_data, colWidths=col_widths, rowHeights=row_heights)
        
        tbl_style = TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), self.font_name),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
        rl_table.setStyle(tbl_style)

        # 绘制位置计算
        # ReportLab wrap 返回建议宽高
        t_w, t_h = rl_table.wrap(w, h)
        
        # PPT 表格坐标是左上角 (x, y)，高度 h
        # PDF 下方坐标是 top_y - actual_height
        top_y = y + h 
        # 这里 t_h 即表格总高，通常应该等于 sum(row_heights)
        draw_y = top_y - t_h 
        
        rl_table.drawOn(c, x, draw_y)

    # --- 辅助函数 ---

    def _get_solid_fill_color(self, fill_obj):
        """
        通用获取颜色函数：支持 SolidFill, Line Properties 等
        """
        try:
            # case 1: 直接是 RGB 颜色
            if fill_obj.type == MSO_COLOR_TYPE.RGB: # 【修复点】使用 MSO_COLOR_TYPE
                return self._rgb_to_color(fill_obj.rgb)
            # case 2: 是 fore_color 属性 (常见于 Shape.fill)
            elif hasattr(fill_obj, 'fore_color'):
                if fill_obj.fore_color.type == MSO_COLOR_TYPE.RGB: # 【修复点】
                    return self._rgb_to_color(fill_obj.fore_color.rgb)
            # case 3: 纯色填充特定检查 (SolidFill)
            elif hasattr(fill_obj, 'solid'):
                 if fill_obj.fore_color.type == MSO_COLOR_TYPE.RGB: # 【修复点】
                    return self._rgb_to_color(fill_obj.fore_color.rgb)
        except Exception:
            pass
        return None

    def _rgb_to_color(self, rgb):
        try:
            return colors.Color(rgb[0]/255.0, rgb[1]/255.0, rgb[2]/255.0)
        except:
            return colors.black