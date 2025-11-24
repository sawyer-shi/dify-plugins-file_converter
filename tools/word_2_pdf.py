import os
import tempfile
import io
import re
import uuid
from typing import Any, Dict, List, Tuple, Optional
from collections.abc import Generator
from collections import defaultdict

# Dify Plugin Imports
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

# --- 依赖库导入 ---
try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.oxml.ns import qn
    
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    # 修正：只导入存在的单位对象
    from reportlab.lib.units import cm, mm
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
    )
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
except ImportError as e:
    raise ImportError(f"Environment Error: {e}. Please ensure requirements.txt is installed.")

# --- 辅助工具 ---
def int_to_chinese(n):
    """数字转中文，用于还原中文列表"""
    chars = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    if n < 10: return chars[n]
    if n < 20: return "十" + (chars[n % 10] if n % 10 != 0 else "")
    return str(n)

class BookmarkParagraph(Paragraph):
    """
    自定义段落组件：
    除了显示文字，还能向 PDF 注册一个书签（Outline）。
    包含【运行时容错】，防止因层级结构错误导致整个PDF生成失败。
    """
    def __init__(self, text, style, level=0):
        super().__init__(text, style)
        self.bookmark_level = level
        self.key = str(uuid.uuid4())

    def draw(self):
        super().draw()
        # 获取当前渲染时的画布对象
        canv = self.canv
        # 添加锚点
        canv.bookmarkPage(self.key)
        
        plaintext = self.getPlainText()
        if plaintext:
            try:
                # 【容错修复】在这里捕获 addOutlineEntry 可能抛出的层级错误
                # title 截取前50字符防止过长
                canv.addOutlineEntry(plaintext[:50], self.key, self.bookmark_level, False)
            except Exception:
                # 如果因为层级结构依然不合法报错，静默失败，不影响PDF生成
                pass

class DocxNumberingEngine:
    """解析 numbering.xml 确保序号 1:1 还原"""
    def __init__(self, doc):
        self.doc = doc
        self.num_dict = {} 
        self.abstract_dict = {} 
        self.counters = {}      
        self._parse_numbering_xml()

    def _parse_numbering_xml(self):
        try:
            numbering_part = self.doc.part.numbering_part
            if not numbering_part: return
            element = numbering_part.element
            for abstract_num in element.findall(qn('w:abstractNum')):
                abs_id = abstract_num.get(qn('w:abstractNumId'))
                levels = {}
                for lvl in abstract_num.findall(qn('w:lvl')):
                    ilvl = int(lvl.get(qn('w:ilvl')))
                    num_fmt = "decimal"
                    fmt_node = lvl.find(qn('w:numFmt'))
                    if fmt_node is not None: num_fmt = fmt_node.get(qn('w:val'))
                    lvl_text = "%1."
                    txt_node = lvl.find(qn('w:lvlText'))
                    if txt_node is not None: lvl_text = txt_node.get(qn('w:val'))
                    levels[ilvl] = (num_fmt, lvl_text)
                self.abstract_dict[abs_id] = levels
            for num in element.findall(qn('w:num')):
                num_id = num.get(qn('w:numId'))
                abs_ref = num.find(qn('w:abstractNumId'))
                if abs_ref is not None:
                    self.num_dict[num_id] = abs_ref.get(qn('w:val'))
        except Exception:
            pass

    def get_numbering_text(self, paragraph) -> str:
        """获取精准的序号字符串"""
        text = paragraph.text.strip()
        if not text: return "" 
        try:
            pPr = paragraph._element.pPr
            if pPr is None or pPr.numPr is None: return ""
            
            num_id_node = pPr.numPr.find(qn('w:numId'))
            if num_id_node is None: return ""
            num_id = num_id_node.get(qn('w:val'))
            
            ilvl_node = pPr.numPr.find(qn('w:ilvl'))
            ilvl = int(ilvl_node.get(qn('w:val'))) if ilvl_node is not None else 0
            
            abstract_id = self.num_dict.get(num_id)
            if not abstract_id: return ""
            level_def = self.abstract_dict.get(abstract_id, {}).get(ilvl)
            if not level_def: return ""
            
            num_fmt, lvl_text = level_def
            counter_key = (num_id, ilvl)
            if counter_key not in self.counters: self.counters[counter_key] = 0
            self.counters[counter_key] += 1
            val = self.counters[counter_key]
            
            if num_fmt == 'bullet': return "• "
            if 'chinese' in num_fmt.lower(): return lvl_text.replace(f'%{ilvl+1}', int_to_chinese(val))
            return lvl_text.replace(f'%{ilvl+1}', str(val)) + " "
        except Exception:
            return ""

class WordToPdfTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        file = tool_parameters.get("input_file")
        if not file or not file.extension.lower().endswith('.docx'):
            yield self.create_text_message("Error: Please upload a .docx file.")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # --- 1. 文件名处理 ---
                original_name = file.filename
                safe_name = re.sub(r'[\\/*?:"<>|]', "", original_name)
                if not safe_name or len(safe_name) < 2: 
                    safe_name = "document.docx"
                if not safe_name.lower().endswith('.docx'): 
                    safe_name += '.docx'
                
                input_path = os.path.join(temp_dir, safe_name)
                with open(input_path, 'wb') as f:
                    f.write(file.blob)

                # 执行转换
                pdf_content, msg = self._convert_to_pdf(input_path, temp_dir)

                if pdf_content:
                    output_filename = os.path.splitext(safe_name)[0] + ".pdf"
                    yield self.create_json_message({
                        "status": "success",
                        "source_file": safe_name,
                        "output_file": output_filename
                    })
                    yield self.create_blob_message(
                        blob=pdf_content,
                        meta={
                            "filename": output_filename, 
                            "mime_type": "application/pdf"
                        }
                    )
                else:
                    yield self.create_text_message(f"Conversion Failed: {msg}")

            except Exception as e:
                import traceback
                yield self.create_text_message(f"Error: {str(e)}\n{traceback.format_exc()}")

    def _register_fonts(self):
        """注册字体，优先查找本地中文字体"""
        font_name = "STSong-Light"
        fonts = {"normal": "STSong-Light", "bold": "STSong-Light"}
        
        candidates = [
            ("msyh.ttf", "MicrosoftYaHei"),
            ("simsun.ttc", "SimSun"),     
            ("simhei.ttf", "SimHei"),     
        ]
        
        search_dirs = [
            os.path.join(os.path.dirname(__file__), "fonts"), 
            "/usr/share/fonts/truetype", 
            "C:\\Windows\\Fonts"
        ]

        for filename, alias in candidates:
            for d in search_dirs:
                path = os.path.join(d, filename)
                if os.path.exists(path):
                    try:
                        pdfmetrics.registerFont(TTFont(alias, path))
                        if alias == "SimSun": fonts["normal"] = alias
                        if alias == "SimHei": fonts["bold"] = alias
                        if alias == "MicrosoftYaHei": fonts["normal"] = alias
                        break
                    except Exception: 
                        continue
        try:
            pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        except Exception: 
            pass
        return fonts

    def _get_col_widths(self, docx_table, total_width_cm=17):
        try:
            tblGrid = docx_table._tbl.tblGrid
            gridCols = tblGrid.gridCol_lst
            if gridCols:
                widths = [int(col.w) for col in gridCols]
                total = sum(widths)
                if total > 0: 
                    return [(w/total)*total_width_cm*cm for w in widths]
        except Exception: 
            pass
        return None 

    def _convert_to_pdf(self, input_path, temp_dir):
        doc = Document(input_path)
        out_path = os.path.join(temp_dir, "temp_out.pdf")
        
        numbering_engine = DocxNumberingEngine(doc)
        
        doc_layout = SimpleDocTemplate(
            out_path, pagesize=A4,
            leftMargin=2*cm, rightMargin=2*cm, topMargin=2.54*cm, bottomMargin=2.54*cm
        )
        
        font_map = self._register_fonts()
        base_font = font_map["normal"]
        bold_font = font_map["bold"] if font_map["bold"] != base_font else base_font
        
        styles = getSampleStyleSheet()
        
        # --- 样式 ---
        style_norm = ParagraphStyle(
            'MyNormal', parent=styles['Normal'],
            fontName=base_font, fontSize=10.5, leading=16,
            wordWrap='CJK', spaceAfter=6, alignment=TA_JUSTIFY
        )
        style_h1 = ParagraphStyle(
            'MyH1', parent=style_norm,
            fontName=bold_font, fontSize=16, leading=22,
            spaceBefore=18, spaceAfter=12, keepWithNext=True
        )
        style_h2 = ParagraphStyle(
            'MyH2', parent=style_norm,
            fontName=bold_font, fontSize=14, leading=20,
            spaceBefore=12, spaceAfter=6, keepWithNext=True
        )
        style_h3 = ParagraphStyle(
            'MyH3', parent=style_norm,
            fontName=bold_font, fontSize=12, leading=18,
            spaceBefore=6, spaceAfter=6, keepWithNext=True
        )

        story = []
        img_map = {}
        try:
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref: img_map[rel.rId] = rel.target_part.blob
        except Exception: pass

        # 【修正】书签层级追踪器，初始为-1（空）
        last_outline_level = -1

        body = doc.element.body
        for child in body.iterchildren():
            
            if isinstance(child, CT_P):
                para = DocxParagraph(child, doc)
                text = para.text.strip()
                
                has_img = False
                for blip in child.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                    rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rid in img_map:
                        try:
                            img_stm = io.BytesIO(img_map[rid])
                            img = ImageReader(img_stm)
                            w, h = img.getSize()
                            if w > 16 * cm: 
                                h = h * (16 * cm / w); w = 16 * cm
                            story.append(RLImage(img_stm, width=w, height=h))
                            has_img = True
                        except Exception: pass
                
                if not text and not has_img: continue

                prefix = numbering_engine.get_numbering_text(para)
                full_text = prefix + text
                if not full_text: continue

                style_name = para.style.name.lower() if para.style else ""
                
                use_style = style_norm
                outline_level = None 
                
                if 'heading 1' in style_name:
                    use_style = style_h1; outline_level = 0
                elif 'heading 2' in style_name:
                    use_style = style_h2; outline_level = 1
                elif 'heading 3' in style_name:
                    use_style = style_h3; outline_level = 2
                elif 'title' in style_name:
                    use_style = style_h1; outline_level = 0
                
                safe_text = full_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                
                if outline_level is not None:
                    # 【逻辑修正】防止跳级 (例如从 -1 跳到 1, 或从 0 跳到 2)
                    # 规则: 新层级最多只能比上一级大 1
                    if outline_level > last_outline_level + 1:
                        outline_level = last_outline_level + 1
                    
                    # 更新追踪器
                    last_outline_level = outline_level
                    
                    p = BookmarkParagraph(safe_text, use_style, level=outline_level)
                else:
                    align = para.alignment
                    if align == 1: p = Paragraph(safe_text, ParagraphStyle('C', parent=use_style, alignment=TA_CENTER))
                    elif align == 2: p = Paragraph(safe_text, ParagraphStyle('R', parent=use_style, alignment=TA_RIGHT))
                    else: p = Paragraph(safe_text, use_style)
                        
                story.append(p)

            elif isinstance(child, CT_Tbl):
                table = DocxTable(child, doc)
                col_widths = self._get_col_widths(table)
                
                rows_data = []
                max_c = len(table.columns)
                if max_c == 0: continue

                for row in table.rows:
                    r_data = []
                    for cell in row.cells:
                        ctext = cell.text.strip().replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        p = Paragraph(ctext, ParagraphStyle('TC', parent=style_norm, fontSize=9, leading=12))
                        r_data.append(p)
                    while len(r_data) < max_c: r_data.append("")
                    rows_data.append(r_data)

                if not rows_data: continue
                
                t = Table(rows_data, colWidths=col_widths)
                t.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('FONTNAME', (0, 0), (-1, -1), base_font),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
                ]))
                story.append(t)
                story.append(Spacer(1, 12))

        try:
            doc_layout.build(story)
            if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                with open(out_path, 'rb') as f:
                    return f.read(), "Success"
            else:
                return None, "PDF generated but empty."
        except Exception as e:
            import traceback
            return None, f"Build Error: {str(e)}"