#!/usr/bin/env python3
"""
DOCX Multi-Industry Document Format Normalizer
Processes .docx files to normalize formatting according to industry-specific standards.

Usage:
    python3 docx_formatter.py --config <config.json>
    python3 docx_formatter.py --input input.docx --template templates/government_document.json --output output.docx
"""

import argparse
import json
import os
import re
import sys
import copy
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Mm, Cm, Inches, Emu, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("ERROR: python-docx is not installed. Run: pip install python-docx")
    sys.exit(1)


class DocxFormatter:
    """Main processing engine for docx format normalization."""

    def __init__(self, config):
        self.config = config
        self.template = config.get("template", {})
        self.modifications = []
        self.doc = None

    @property
    def industry_name(self):
        return self.template.get("industry_name", "Unknown")

    def load_document(self, input_path):
        try:
            self.doc = Document(input_path)
        except Exception as e:
            raise RuntimeError(f"Failed to parse docx file: {e}")

    def run(self, input_path, output_path):
        self.load_document(input_path)
        self._apply_page_setup()
        self._apply_body_text()
        self._apply_headings()
        self._apply_hierarchy_fonts()
        self._apply_header_footer()
        self._apply_tables()
        self._apply_force_clear()
        self._apply_page_zones_if_needed()
        self.doc.save(output_path)
        return self.modifications

    # ── Page Setup ──────────────────────────────────────────

    def _apply_page_setup(self):
        page_setup = self.template.get("page_setup")
        if not page_setup:
            return

        for section in self.doc.sections:
            # Paper size
            if page_setup.get("paper_size") == "A4":
                section.page_width = Mm(210)
                section.page_height = Mm(297)
                self._record("页面设置", "纸张大小", "非A4或默认", "A4 (210×297mm)")

            # Orientation
            if page_setup.get("orientation") == "portrait":
                section.orientation = WD_ORIENT.PORTRAIT

            # Margins
            margins = page_setup.get("margins", {})
            margin_map = {
                "top_mm": ("top_margin", Mm),
                "bottom_mm": ("bottom_margin", Mm),
                "left_mm": ("left_margin", Mm),
                "right_mm": ("right_margin", Mm),
            }
            for key, (attr, unit) in margin_map.items():
                if key in margins:
                    new_val = margins[key]
                    old_val = getattr(section, attr)
                    old_mm = round(old_val / 36000, 1) if old_val else 0
                    if abs(old_mm - new_val) > 0.5:
                        setattr(section, attr, unit(new_val))
                        self._record("页面设置", f"边距({key})",
                                     f"{old_mm}mm", f"{new_val}mm")

            # Lines per page & chars per line (government/court docs)
            lines_per_page = page_setup.get("lines_per_page")
            chars_per_line = page_setup.get("chars_per_line")
            if lines_per_page and chars_per_line:
                self._set_grid_lines(section, lines_per_page, chars_per_line)

    def _set_grid_lines(self, section, lines_per_page, chars_per_line):
        """Set document grid for fixed lines/chars per page (公文/法院文书)."""
        sectPr = section._sectPr
        docGrid = sectPr.find(qn('w:docGrid'))
        if docGrid is None:
            docGrid = OxmlElement('w:docGrid')
            sectPr.append(docGrid)
        # linePitch: 312 twips ≈ 1 line height for 3号 font
        line_pitch = 312
        docGrid.set(qn('w:type'), 'lines')
        docGrid.set(qn('w:linePitch'), str(line_pitch))
        self._record("页面设置", "文档网格",
                     "默认", f"{lines_per_page}行×{chars_per_line}字/页")

    # ── Body Text ────────────────────────────────────────────

    def _apply_body_text(self):
        body = self.template.get("body_text")
        if not body:
            return

        default_body = self.template.get("default_body_text", body)

        for para in self.doc.paragraphs:
            if self._is_heading_paragraph(para):
                continue
            self._format_paragraph(para, default_body)

    def _format_paragraph(self, para, fmt):
        font_cn = fmt.get("font_cn", "宋体")
        font_en = fmt.get("font_en", "Times New Roman")
        font_size = fmt.get("font_size_pt")
        line_spacing_rule = fmt.get("line_spacing_rule", "1.5x")
        first_indent = fmt.get("first_line_indent_chars", 0)
        alignment = fmt.get("alignment", "left")
        text_color = fmt.get("text_color", "000000")

        changed = False
        changes = []

        # Paragraph format: alignment
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if alignment in align_map:
            old_align = para.alignment
            new_align = align_map[alignment]
            if old_align != new_align:
                para.alignment = new_align
                changes.append(f"对齐: {old_align}→{alignment}")

        # Line spacing
        pf = para.paragraph_format
        if line_spacing_rule == "1.5x":
            pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            changes.append("行距→1.5倍")
        elif line_spacing_rule == "single":
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            changes.append("行距→单倍")
        elif line_spacing_rule == "fixed":
            exact_pt = fmt.get("line_spacing_exact_pt", 28.8)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(exact_pt)
            changes.append(f"行距→固定{exact_pt}pt")

        # First line indent
        if first_indent > 0:
            pf.first_line_indent = Pt(font_size * first_indent)
            changes.append(f"首行缩进→{first_indent}字符")

        # Space before/after
        space_before = fmt.get("space_before_lines")
        space_after = fmt.get("space_after_lines")
        if space_before is not None:
            pf.space_before = Pt(font_size * space_before)
        if space_after is not None:
            pf.space_after = Pt(font_size * space_after)

        # Run-level formatting
        for run in para.runs:
            run_changes = self._format_run(run, font_cn, font_en, font_size, text_color, fmt)
            changes.extend(run_changes)

        if changes:
            preview = para.text[:30] + ("..." if len(para.text) > 30 else "")
            self._record("正文段落", preview, "—", "; ".join(changes[:3]))

    def _format_run(self, run, font_cn, font_en, font_size, text_color, fmt=None):
        changes = []

        # Font name (Chinese)
        old_font = run.font.name
        if font_cn and run.font.name != font_cn:
            run.font.name = font_cn
            r = run._element
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:eastAsia'), font_cn)
            rFonts.set(qn('w:ascii'), font_en or font_cn)
            rFonts.set(qn('w:hAnsi'), font_en or font_cn)
            changes.append(f"字体→{font_cn}")

        # Font size
        if font_size and (run.font.size is None or run.font.size != Pt(font_size)):
            run.font.size = Pt(font_size)
            changes.append(f"字号→{font_size}pt")

        # Text color
        if text_color:
            color = RGBColor.from_string(text_color)
            if run.font.color is None or run.font.color.rgb != color:
                run.font.color.rgb = color
                changes.append("颜色→黑色")

        # Char spacing (for court documents)
        if fmt and fmt.get("char_spacing_pt"):
            r = run._element
            rPr = r.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                r.insert(0, rPr)
            spacing = rPr.find(qn('w:spacing'))
            if spacing is None:
                spacing = OxmlElement('w:spacing')
                rPr.append(spacing)
            spacing.set(qn('w:val'), str(int(fmt["char_spacing_pt"] * 20)))

        return changes

    # ── Headings ─────────────────────────────────────────────

    def _is_heading_paragraph(self, para):
        style_name = (para.style.name or "").lower() if para.style else ""
        return "heading" in style_name or "title" in style_name

    def _apply_headings(self):
        headings = self.template.get("headings", [])
        if not headings:
            return

        for para in self.doc.paragraphs:
            for h in headings:
                if self._match_heading(para, h):
                    self._format_heading(para, h)
                    break

    def _match_heading(self, para, h):
        style_name = (para.style.name or "").lower() if para.style else ""
        level = h.get("level", 0)
        if level == 0:
            return "title" in style_name
        return f"heading {level}" in style_name

    def _format_heading(self, para, h):
        font_cn = h.get("font", h.get("font_cn", "宋体"))
        font_en = h.get("font_en", "Times New Roman")
        font_size = h.get("font_size_pt", 14)
        bold = h.get("bold", False)
        alignment = h.get("alignment", "center")

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }
        if alignment in align_map:
            para.alignment = align_map[alignment]

        for run in para.runs:
            self._format_run(run, font_cn, font_en, font_size, "000000")
            run.font.bold = bold

        # Space before/after
        if h.get("space_before_lines") is not None:
            para.paragraph_format.space_before = Pt(font_size * h["space_before_lines"])
        if h.get("space_after_lines") is not None:
            para.paragraph_format.space_after = Pt(font_size * h["space_after_lines"])

        preview = para.text[:30] + ("..." if len(para.text) > 30 else "")
        self._record("标题", preview, "—",
                     f"{font_cn} {font_size}pt {'加粗' if bold else ''} {alignment}")

    # ── Hierarchy Fonts ─────────────────────────────────────

    def _apply_hierarchy_fonts(self):
        hier_fonts = self.template.get("hierarchy_fonts", [])
        if not hier_fonts:
            return

        for para in self.doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            for h in hier_fonts:
                pattern = h.get("prefix_pattern", "")
                if pattern and re.search(pattern, text):
                    font_cn = h.get("font", "仿宋")
                    font_size = h.get("font_size_pt", 16)
                    bold = h.get("bold", False)
                    for run in para.runs:
                        self._format_run(run, font_cn, "Times New Roman", font_size, "000000")
                        run.font.bold = bold
                    preview = text[:30] + ("..." if len(text) > 30 else "")
                    self._record("层次序号", preview, "—",
                                 f"{font_cn} {font_size}pt")
                    break

    # ── Header / Footer ──────────────────────────────────────

    def _apply_header_footer(self):
        hf = self.template.get("header_footer")
        if not hf:
            return

        for section in self.doc.sections:
            # Remove headers/footers if required
            if not hf.get("has_header", True):
                header = section.header
                for para in header.paragraphs:
                    for run in para.runs:
                        run.text = ""
                self._record("页眉", "清除页眉内容", "有内容", "已清空")

            if not hf.get("has_footer", True):
                footer = section.footer
                for para in footer.paragraphs:
                    for run in para.runs:
                        run.text = ""
                self._record("页脚", "清除页脚内容", "有内容", "已清空")

            if not hf.get("has_page_number", True):
                self._remove_page_numbers(section)
                self._record("页码", "清除页码", "有页码", "已移除")

            # Add page number if required (government/court docs)
            if hf.get("footer_type") == "page_number" and hf.get("has_page_number", True):
                self._set_page_number_format(section, hf)

    def _remove_page_numbers(self, section):
        for part in [section.header, section.footer]:
            for para in part.paragraphs:
                for run in para.runs:
                    fld_elems = run._element.findall(qn('w:fldSimple'))
                    for fld in fld_elems:
                        if 'PAGE' in fld.get(qn('w:instr'), ''):
                            fld.getparent().remove(fld)

    def _set_page_number_format(self, section, hf):
        font_name = hf.get("page_number_font", "宋体")
        font_size = hf.get("page_number_size_pt", 14)
        odd_align = hf.get("page_number_odd_align", "right")
        even_align = hf.get("page_number_even_align", "left")

        footer = section.footer
        if not footer.paragraphs or not footer.paragraphs[0].text:
            para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            run = para.add_run()
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = ' PAGE '
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'end')
            run._element.append(fldChar1)
            run._element.append(instrText)
            run._element.append(fldChar2)
            run.font.name = font_name
            run.font.size = Pt(font_size)

        align_map = {
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
        }
        if odd_align in align_map:
            footer.paragraphs[0].alignment = align_map[odd_align]

        self._record("页码", "设置页码格式", "—",
                     f"{font_name} {font_size}pt {odd_align}对齐")

    # ── Tables ───────────────────────────────────────────────

    def _apply_tables(self):
        table_cfg = self.template.get("tables")
        if not table_cfg:
            return

        for table in self.doc.tables:
            border_width = table_cfg.get("border_width_pt", 0.5)
            self._set_table_borders(table, border_width)

            if table_cfg.get("table_alignment") == "center":
                table.alignment = 1  # WD_TABLE_ALIGNMENT.CENTER

            cell_font_cn = table_cfg.get("cell_font_cn", "宋体")
            cell_font_en = table_cfg.get("cell_font_en", "Times New Roman")
            cell_font_size = table_cfg.get("cell_font_size_pt")

            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if cell_font_cn:
                                self._format_run(run, cell_font_cn, cell_font_en,
                                                 cell_font_size, "000000")

        self._record("表格", f"共{len(self.doc.tables)}个表格", "—",
                     f"边框{table_cfg.get('border_width_pt', 0.5)}磅 居中")

    def _set_table_borders(self, table, width_pt):
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        borders = tblPr.find(qn('w:tblBorders'))
        if borders is None:
            borders = OxmlElement('w:tblBorders')
            tblPr.append(borders)
        for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            elem = borders.find(qn(f'w:{edge}'))
            if elem is None:
                elem = OxmlElement(f'w:{edge}')
                borders.append(elem)
            elem.set(qn('w:val'), 'single')
            elem.set(qn('w:sz'), str(int(width_pt * 8)))
            elem.set(qn('w:color'), '000000')

    # ── Force Clear (Dark-bid mode) ─────────────────────────

    def _apply_force_clear(self):
        clear_rules = self.template.get("force_clear", [])
        if not clear_rules:
            return

        for rule in clear_rules:
            if rule == "bold":
                self._clear_bold()
            elif rule == "italic":
                self._clear_italic()
            elif rule == "underline":
                self._clear_underline()
            elif rule == "strikethrough":
                self._clear_strikethrough()
            elif rule in ("colored_text", "colored_background"):
                self._clear_colors()
            elif rule == "headers":
                self._clear_headers()
            elif rule == "footers":
                self._clear_footers()
            elif rule == "page_numbers":
                self._clear_page_numbers()
            elif rule == "table_of_contents":
                self._clear_toc()
            elif rule == "hyperlink_coloring":
                self._clear_hyperlink_colors()
            elif rule == "custom_char_spacing":
                self._clear_char_spacing()
            elif rule == "custom_char_scale":
                self._clear_char_scale()

    def _clear_bold(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                if run.font.bold:
                    run.font.bold = False
                    count += 1
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.font.bold:
                                run.font.bold = False
                                count += 1
        if count:
            self._record("强制清除", "加粗", f"{count}处加粗", "全部清除")

    def _clear_italic(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                if run.font.italic:
                    run.font.italic = False
                    count += 1
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.font.italic:
                                run.font.italic = False
                                count += 1
        if count:
            self._record("强制清除", "倾斜", f"{count}处倾斜", "全部清除")

    def _clear_underline(self):
        count = 0
        from docx.enum.text import WD_UNDERLINE
        for para in self.doc.paragraphs:
            for run in para.runs:
                if run.font.underline and run.font.underline != WD_UNDERLINE.NONE:
                    run.font.underline = WD_UNDERLINE.NONE
                    count += 1
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.font.underline and run.font.underline != WD_UNDERLINE.NONE:
                                run.font.underline = WD_UNDERLINE.NONE
                                count += 1
        if count:
            self._record("强制清除", "下划线", f"{count}处下划线", "全部清除")

    def _clear_strikethrough(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                if run.font.strike:
                    run.font.strike = False
                    count += 1
        if count:
            self._record("强制清除", "删除线", f"{count}处", "全部清除")

    def _clear_colors(self):
        count = 0
        black = RGBColor(0, 0, 0)
        for para in self.doc.paragraphs:
            for run in para.runs:
                if run.font.color and run.font.color.rgb and run.font.color.rgb != black:
                    run.font.color.rgb = black
                    count += 1
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if run.font.color and run.font.color.rgb and run.font.color.rgb != black:
                                run.font.color.rgb = black
                                count += 1
        if count:
            self._record("强制清除", "彩色字体", f"{count}处彩色", "全部改为黑色")

    def _clear_headers(self):
        for section in self.doc.sections:
            header = section.header
            for para in header.paragraphs:
                for run in para.runs:
                    run.text = ""
        self._record("强制清除", "页眉", "有页眉", "全部清空")

    def _clear_footers(self):
        for section in self.doc.sections:
            footer = section.footer
            for para in footer.paragraphs:
                for run in para.runs:
                    run.text = ""
        self._record("强制清除", "页脚", "有页脚", "全部清空")

    def _clear_page_numbers(self):
        for section in self.doc.sections:
            self._remove_page_numbers(section)
        self._record("强制清除", "页码", "有页码", "全部移除")

    def _clear_toc(self):
        count = 0
        for para in self.doc.paragraphs:
            style_name = (para.style.name or "").lower() if para.style else ""
            if "toc" in style_name:
                count += 1
        if count:
            self._record("强制清除", "目录", f"{count}处目录", "已标记需手动删除")

    def _clear_hyperlink_colors(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                r = run._element
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    color = rPr.find(qn('w:color'))
                    if color is not None:
                        val = color.get(qn('w:val'))
                        if val and val.upper() not in ('000000', 'AUTO'):
                            color.set(qn('w:val'), '000000')
                            count += 1
                    underline = rPr.find(qn('w:u'))
                    if underline is not None:
                        underline.set(qn('w:val'), 'none')
        if count:
            self._record("强制清除", "超链接颜色", f"{count}处", "改为黑色无下划线")

    def _clear_char_spacing(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                r = run._element
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    spacing = rPr.find(qn('w:spacing'))
                    if spacing is not None and spacing.get(qn('w:val')):
                        spacing.set(qn('w:val'), '0')
                        count += 1
        if count:
            self._record("强制清除", "字符间距", f"{count}处自定义间距", "恢复标准")

    def _clear_char_scale(self):
        count = 0
        for para in self.doc.paragraphs:
            for run in para.runs:
                r = run._element
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    w = rPr.find(qn('w:w'))
                    if w is not None:
                        val = w.get(qn('w:val'))
                        if val and val != '100':
                            w.set(qn('w:val'), '100')
                            count += 1
        if count:
            self._record("强制清除", "字符缩放", f"{count}处非100%", "恢复100%")

    # ── Page Zones (Testing Reports) ────────────────────────

    def _apply_page_zones_if_needed(self):
        page_zones = self.template.get("page_zones")
        if not page_zones:
            return

        # Testing report: detect zone by keyword and apply per-zone formatting
        current_zone = None
        para_index = 0
        total_paras = len(self.doc.paragraphs)

        for i, para in enumerate(self.doc.paragraphs):
            text = para.text.strip()
            zone = self._detect_page_zone(text, i, total_paras, page_zones, current_zone)

            if zone:
                current_zone = zone.get("zone_id")
                elements = zone.get("elements", [])
                for elem in elements:
                    if self._match_zone_element(para, elem, i):
                        self._apply_zone_element(para, elem)
                        break

    def _detect_page_zone(self, text, para_idx, total_paras, page_zones, current_zone):
        for zone in page_zones:
            keywords = zone.get("detection_keywords", [])
            position = zone.get("detection_position", "")

            if position == "first_page" and para_idx < 10:
                for kw in keywords:
                    if kw in text:
                        return zone

            if not position:
                for kw in keywords:
                    if kw in text:
                        return zone

        return None

    def _match_zone_element(self, para, elem, idx):
        content_hint = elem.get("content_hint", "")
        if content_hint and content_hint in para.text:
            return True
        return False

    def _apply_zone_element(self, para, elem):
        font = elem.get("font", "仿宋")
        font_size = elem.get("font_size_pt", 12)
        bold = elem.get("bold", False)
        alignment = elem.get("alignment")

        if alignment:
            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
            }
            if alignment in align_map:
                para.alignment = align_map[alignment]

        for run in para.runs:
            self._format_run(run, font, "Times New Roman", font_size, "000000")
            run.font.bold = bold

        preview = para.text[:30] + ("..." if len(para.text) > 30 else "")
        self._record("页面分区", preview, "—",
                     f"{font} {font_size}pt {'加粗' if bold else ''}")

    # ── Report Generation ────────────────────────────────────

    def _record(self, location, detail, before, after):
        self.modifications.append({
            "location": location,
            "detail": detail,
            "before": before,
            "after": after,
        })

    def generate_report(self, output_report_path):
        report = {
            "industry": self.template.get("industry_name", "Unknown"),
            "standard": self.template.get("standard", ""),
            "total_modifications": len(self.modifications),
            "modifications": self.modifications,
        }
        with open(output_report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        return report


def load_template(template_path):
    with open(template_path, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    parser = argparse.ArgumentParser(
        description="DOCX Multi-Industry Document Format Normalizer"
    )
    parser.add_argument("--config", help="Path to JSON config file (full config)")
    parser.add_argument("--input", help="Input docx file path")
    parser.add_argument("--template", help="Template JSON file path")
    parser.add_argument("--output", help="Output docx file path")
    parser.add_argument("--report", help="Output report JSON file path")
    args = parser.parse_args()

    # Load configuration
    if args.config:
        with open(args.config, "r", encoding="utf-8") as f:
            config = json.load(f)
        input_path = config["input_file"]
        output_path = config.get("output_file", "output_normalized.docx")
        report_path = config.get("report_file", "modification_report.json")
        template = load_template(config["template_path"]) if "template_path" in config else config.get("template", {})
    else:
        if not args.input or not args.template:
            parser.error("--input and --template are required (or use --config)")
        input_path = args.input
        output_path = args.output or "output_normalized.docx"
        report_path = args.report or "modification_report.json"
        template = load_template(args.template)

    # Process
    formatter = DocxFormatter({"template": template})
    try:
        formatter.load_document(input_path)
    except RuntimeError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print(f"[INFO] Industry: {formatter.industry_name}")
    print(f"[INFO] Input: {input_path}")
    print(f"[INFO] Output: {output_path}")

    modifications = formatter.run(input_path, output_path)
    report = formatter.generate_report(report_path)

    print(f"\n[DONE] Processing complete.")
    print(f"  - Total modifications: {len(modifications)}")
    print(f"  - Output file: {output_path}")
    print(f"  - Report file: {report_path}")
    print(f"\n--- Modification Summary ---")
    for m in modifications:
        print(f"  [{m['location']}] {m['detail']}: {m['before']} → {m['after']}")


if __name__ == "__main__":
    main()
