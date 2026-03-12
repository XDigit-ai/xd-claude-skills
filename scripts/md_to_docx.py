#!/usr/bin/env python3
"""
Convert Markdown to branded DOCX with XD.AI template styling.
v3: improved tables, optional cover page, optional TOC, H1 page breaks.
"""

import argparse
import datetime
import os
import re
import shutil
import struct
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SKILL_DIR = Path(__file__).parent.parent
TEMPLATE_PATH = SKILL_DIR / "assets" / "template.dotx"
LOGO_PATH = SKILL_DIR / "assets" / "logo.png"

NAMESPACES = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'wp14':'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

EXTRA_NS = [
    ('wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'),
    ('mc',  'http://schemas.openxmlformats.org/markup-compatibility/2006'),
    ('o',   'urn:schemas-microsoft-com:office:office'),
    ('m',   'http://schemas.openxmlformats.org/officeDocument/2006/math'),
    ('v',   'urn:schemas-microsoft-com:vml'),
    ('w10', 'urn:schemas-microsoft-com:office:word'),
    ('w15', 'http://schemas.microsoft.com/office/word/2012/wordml'),
    ('wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'),
    ('wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk'),
    ('wne', 'http://schemas.microsoft.com/office/word/2006/wordml'),
    ('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'),
]
for prefix, uri in EXTRA_NS:
    ET.register_namespace(prefix, uri)


class MarkdownToDocx:
    """Convert Markdown to branded DOCX."""

    def __init__(self, template_path=None, cover=False, toc=False):
        self.template_path = template_path or TEMPLATE_PATH
        self.cover = cover
        self.toc = toc
        self.temp_dir = None
        self._h1_seen = False

    def convert(self, md_content: str, output_path: str):
        self._h1_seen = False
        self.temp_dir = tempfile.mkdtemp()
        template_dir = os.path.join(self.temp_dir, 'template')

        with zipfile.ZipFile(self.template_path, 'r') as zf:
            zf.extractall(template_dir)

        paragraphs = self._parse_markdown(md_content)
        self._add_footer(template_dir)
        self._update_document(template_dir, paragraphs)
        self._create_docx(template_dir, output_path)

        shutil.rmtree(self.temp_dir)
        return output_path

    # ------------------------------------------------------------------ #
    # Markdown parsing                                                     #
    # ------------------------------------------------------------------ #

    def _parse_markdown(self, content: str) -> list:
        paragraphs = []
        lines = content.split('\n')
        i = 0
        in_code_block = False
        code_block_content = []
        code_lang = ''

        while i < len(lines):
            line = lines[i]

            if line.startswith('```'):
                if not in_code_block:
                    in_code_block = True
                    code_lang = line[3:].strip()
                    code_block_content = []
                else:
                    in_code_block = False
                    paragraphs.append({
                        'type': 'code_block',
                        'content': '\n'.join(code_block_content),
                        'language': code_lang,
                    })
                i += 1
                continue

            if in_code_block:
                code_block_content.append(line)
                i += 1
                continue

            if line.startswith('#'):
                match = re.match(r'^(#{1,6})\s+(.+)$', line)
                if match:
                    paragraphs.append({
                        'type': 'heading',
                        'level': len(match.group(1)),
                        'content': match.group(2),
                    })
                    i += 1
                    continue

            if re.match(r'^(-{3,}|\*{3,}|_{3,})$', line.strip()):
                paragraphs.append({'type': 'hr'})
                i += 1
                continue

            if re.match(r'^(\s*)[-*+]\s+', line):
                match = re.match(r'^(\s*)([-*+])\s+(.+)$', line)
                if match:
                    paragraphs.append({
                        'type': 'bullet',
                        'level': len(match.group(1)) // 2,
                        'content': match.group(3),
                    })
                    i += 1
                    continue

            if re.match(r'^(\s*)\d+\.\s+', line):
                match = re.match(r'^(\s*)(\d+)\.\s+(.+)$', line)
                if match:
                    paragraphs.append({
                        'type': 'numbered',
                        'level': len(match.group(1)) // 3,
                        'number': int(match.group(2)),
                        'content': match.group(3),
                    })
                    i += 1
                    continue

            if line.startswith('>'):
                paragraphs.append({'type': 'quote', 'content': line.lstrip('>').strip()})
                i += 1
                continue

            # Table: current line has | and next line is a separator row
            if '|' in line and i + 1 < len(lines) and re.match(r'^\s*\|?\s*[-:]+[-| :]*\s*$', lines[i + 1]):
                table_rows = []
                j = i
                while j < len(lines) and '|' in lines[j]:
                    row_line = lines[j].strip().strip('|')
                    cells = [c.strip() for c in row_line.split('|')]
                    table_rows.append(cells)
                    j += 1
                if len(table_rows) >= 2 and re.match(r'^[\s\-:]+$', table_rows[1][0].replace('|', '')):
                    header = table_rows[0]
                    data_rows = table_rows[2:]
                else:
                    header = None
                    data_rows = table_rows
                paragraphs.append({'type': 'table', 'header': header, 'rows': data_rows})
                i = j
                continue

            if not line.strip():
                i += 1
                continue

            paragraphs.append({'type': 'paragraph', 'content': line})
            i += 1

        return paragraphs

    # ------------------------------------------------------------------ #
    # Inline formatting                                                    #
    # ------------------------------------------------------------------ #

    def _create_run_with_formatting(self, text: str, parent: ET.Element):
        patterns = [
            (r'\*\*\*(.+?)\*\*\*', 'bold_italic'),
            (r'\*\*(.+?)\*\*',     'bold'),
            (r'__(.+?)__',         'bold'),
            (r'\*(.+?)\*',         'italic'),
            (r'_(.+?)_',           'italic'),
            (r'`(.+?)`',           'code'),
            (r'\[(.+?)\]\((.+?)\)','link'),
        ]
        remaining = text
        while remaining:
            earliest_match = None
            earliest_pos = len(remaining)
            match_type = None
            for pattern, fmt_type in patterns:
                m = re.search(pattern, remaining)
                if m and m.start() < earliest_pos:
                    earliest_match = m
                    earliest_pos = m.start()
                    match_type = fmt_type
            if earliest_match:
                if earliest_pos > 0:
                    self._add_run(parent, remaining[:earliest_pos])
                if match_type == 'bold':
                    self._add_run(parent, earliest_match.group(1), bold=True)
                elif match_type == 'italic':
                    self._add_run(parent, earliest_match.group(1), italic=True)
                elif match_type == 'bold_italic':
                    self._add_run(parent, earliest_match.group(1), bold=True, italic=True)
                elif match_type == 'code':
                    self._add_run(parent, earliest_match.group(1), code=True)
                elif match_type == 'link':
                    self._add_run(parent, earliest_match.group(1), underline=True, color='3313E2')
                remaining = remaining[earliest_match.end():]
            else:
                self._add_run(parent, remaining)
                break

    def _add_run(self, parent: ET.Element, text: str, bold=False, italic=False,
                 code=False, underline=False, color=None):
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        run = ET.SubElement(parent, f'{w}r')
        if bold or italic or code or underline or color:
            rPr = ET.SubElement(run, f'{w}rPr')
            if bold:
                ET.SubElement(rPr, f'{w}b')
            if italic:
                ET.SubElement(rPr, f'{w}i')
            if underline:
                u = ET.SubElement(rPr, f'{w}u')
                u.set(f'{w}val', 'single')
            if color:
                c = ET.SubElement(rPr, f'{w}color')
                c.set(f'{w}val', color)
            if code:
                rFonts = ET.SubElement(rPr, f'{w}rFonts')
                rFonts.set(f'{w}ascii', 'Consolas')
                rFonts.set(f'{w}hAnsi', 'Consolas')
                shd = ET.SubElement(rPr, f'{w}shd')
                shd.set(f'{w}val', 'clear')
                shd.set(f'{w}fill', 'E8E8E8')
        t = ET.SubElement(run, f'{w}t')
        t.text = text
        if text.startswith(' ') or text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    # ------------------------------------------------------------------ #
    # Paragraph creation                                                   #
    # ------------------------------------------------------------------ #

    def _create_paragraph(self, para_data: dict) -> ET.Element:
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        p = ET.Element(f'{w}p')
        pPr = ET.SubElement(p, f'{w}pPr')
        para_type = para_data['type']

        if para_type == 'heading':
            level = para_data['level']
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', f'Heading{level}')
            # Page break before H1 (except the first one)
            if level == 1:
                if self._h1_seen:
                    pbr = ET.SubElement(pPr, f'{w}pageBreakBefore')
                    pbr.set(f'{w}val', 'true')
                else:
                    self._h1_seen = True
            self._create_run_with_formatting(para_data['content'], p)

        elif para_type == 'paragraph':
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', 'Normal')
            self._create_run_with_formatting(para_data['content'], p)

        elif para_type == 'bullet':
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', 'ListParagraph')
            numPr = ET.SubElement(pPr, f'{w}numPr')
            ilvl = ET.SubElement(numPr, f'{w}ilvl')
            ilvl.set(f'{w}val', str(para_data.get('level', 0)))
            numId = ET.SubElement(numPr, f'{w}numId')
            numId.set(f'{w}val', '1')
            self._create_run_with_formatting(para_data['content'], p)

        elif para_type == 'numbered':
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', 'ListParagraph')
            numPr = ET.SubElement(pPr, f'{w}numPr')
            ilvl = ET.SubElement(numPr, f'{w}ilvl')
            ilvl.set(f'{w}val', str(para_data.get('level', 0)))
            numId = ET.SubElement(numPr, f'{w}numId')
            numId.set(f'{w}val', '2')
            self._create_run_with_formatting(para_data['content'], p)

        elif para_type == 'quote':
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', 'Quote')
            self._create_run_with_formatting(para_data['content'], p)

        elif para_type == 'code_block':
            pStyle = ET.SubElement(pPr, f'{w}pStyle')
            pStyle.set(f'{w}val', 'Normal')
            pShd = ET.SubElement(pPr, f'{w}shd')
            pShd.set(f'{w}val', 'clear')
            pShd.set(f'{w}fill', 'E8E8E8')
            spacing = ET.SubElement(pPr, f'{w}spacing')
            spacing.set(f'{w}before', '120')
            spacing.set(f'{w}after', '120')
            run = ET.SubElement(p, f'{w}r')
            rPr = ET.SubElement(run, f'{w}rPr')
            rFonts = ET.SubElement(rPr, f'{w}rFonts')
            rFonts.set(f'{w}ascii', 'Consolas')
            rFonts.set(f'{w}hAnsi', 'Consolas')
            sz = ET.SubElement(rPr, f'{w}sz')
            sz.set(f'{w}val', '20')
            for idx, cl in enumerate(para_data['content'].split('\n')):
                if idx > 0:
                    ET.SubElement(run, f'{w}br')
                t = ET.SubElement(run, f'{w}t')
                t.text = cl
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        elif para_type == 'hr':
            pBdr = ET.SubElement(pPr, f'{w}pBdr')
            bottom = ET.SubElement(pBdr, f'{w}bottom')
            bottom.set(f'{w}val', 'single')
            bottom.set(f'{w}sz', '6')
            bottom.set(f'{w}space', '1')
            bottom.set(f'{w}color', 'E8E8E8')

        return p

    # ------------------------------------------------------------------ #
    # Table cell helpers                                                   #
    # ------------------------------------------------------------------ #

    def _add_table_run(self, parent: ET.Element, text: str,
                       is_header=False, bold=False, italic=False):
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        run = ET.SubElement(parent, f'{w}r')
        rPr = ET.SubElement(run, f'{w}rPr')
        rFonts = ET.SubElement(rPr, f'{w}rFonts')
        rFonts.set(f'{w}ascii', 'Causten')
        rFonts.set(f'{w}hAnsi', 'Causten')
        sz = ET.SubElement(rPr, f'{w}sz')
        sz.set(f'{w}val', '20')
        szCs = ET.SubElement(rPr, f'{w}szCs')
        szCs.set(f'{w}val', '20')
        if bold or is_header:
            ET.SubElement(rPr, f'{w}b')
        if italic:
            ET.SubElement(rPr, f'{w}i')
        if is_header:
            clr = ET.SubElement(rPr, f'{w}color')
            clr.set(f'{w}val', 'FFFFFF')
        t = ET.SubElement(run, f'{w}t')
        t.text = text
        if text.startswith(' ') or text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    def _create_table_cell_runs(self, text: str, parent: ET.Element, is_header=False):
        """Parse inline bold/italic in table cells and create formatted runs."""
        patterns = [
            (r'\*\*\*(.+?)\*\*\*', 'bold_italic'),
            (r'\*\*(.+?)\*\*',     'bold'),
            (r'__(.+?)__',         'bold'),
            (r'\*(.+?)\*',         'italic'),
            (r'_(.+?)_',           'italic'),
        ]
        remaining = text
        while remaining:
            earliest_match = None
            earliest_pos = len(remaining)
            match_type = None
            for pattern, fmt_type in patterns:
                m = re.search(pattern, remaining)
                if m and m.start() < earliest_pos:
                    earliest_match = m
                    earliest_pos = m.start()
                    match_type = fmt_type
            if earliest_match:
                if earliest_pos > 0:
                    self._add_table_run(parent, remaining[:earliest_pos], is_header=is_header)
                bold = match_type in ('bold', 'bold_italic')
                italic = match_type in ('italic', 'bold_italic')
                self._add_table_run(parent, earliest_match.group(1),
                                    is_header=is_header, bold=bold, italic=italic)
                remaining = remaining[earliest_match.end():]
            else:
                self._add_table_run(parent, remaining, is_header=is_header)
                break

    # ------------------------------------------------------------------ #
    # Table creation                                                       #
    # ------------------------------------------------------------------ #

    def _create_table(self, table_data: dict) -> ET.Element:
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        HEADER_BG    = '3313E2'
        ALT_ROW_BG   = 'EEF0FD'
        BORDER_COLOR = 'D0D0D0'
        BORDER_SZ    = '4'
        PAGE_WIDTH   = 9070  # twips — A4 with 25mm margins each side

        def make_border(tag):
            el = ET.Element(f'{w}{tag}')
            el.set(f'{w}val',   'single')
            el.set(f'{w}sz',    BORDER_SZ)
            el.set(f'{w}space', '0')
            el.set(f'{w}color', BORDER_COLOR)
            return el

        header    = table_data.get('header') or []
        data_rows = table_data.get('rows')   or []
        all_rows  = ([header] if header else []) + data_rows
        num_cols  = max((len(r) for r in all_rows), default=1)

        # Proportional column widths based on max content length per column
        col_lengths = []
        for col_idx in range(num_cols):
            max_len = 0
            for row in all_rows:
                if col_idx < len(row):
                    # Strip markdown markers before measuring
                    text = re.sub(r'\*+', '', row[col_idx])
                    max_len = max(max_len, len(text))
            col_lengths.append(max(max_len, 4))
        total_chars = sum(col_lengths)
        min_col = max(900, PAGE_WIDTH // (num_cols * 3))  # floor: ~1/3 of equal share
        col_widths = []
        remaining_w = PAGE_WIDTH
        for length in col_lengths[:-1]:
            cw = max(int(PAGE_WIDTH * length / total_chars), min_col)
            col_widths.append(cw)
            remaining_w -= cw
        col_widths.append(max(remaining_w, min_col))

        tbl = ET.Element(f'{w}tbl')
        tblPr = ET.SubElement(tbl, f'{w}tblPr')
        tblW = ET.SubElement(tblPr, f'{w}tblW')
        tblW.set(f'{w}w', str(PAGE_WIDTH))
        tblW.set(f'{w}type', 'dxa')
        tblBorders = ET.SubElement(tblPr, f'{w}tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            tblBorders.append(make_border(side))
        tblLook = ET.SubElement(tblPr, f'{w}tblLook')
        tblLook.set(f'{w}val', '04A0')

        tblGrid = ET.SubElement(tbl, f'{w}tblGrid')
        for cw in col_widths:
            gc = ET.SubElement(tblGrid, f'{w}gridCol')
            gc.set(f'{w}w', str(cw))

        def add_row(cells, is_header=False, row_index=0):
            tr = ET.SubElement(tbl, f'{w}tr')
            trPr = ET.SubElement(tr, f'{w}trPr')
            if is_header:
                th = ET.SubElement(trPr, f'{w}tblHeader')
                th.set(f'{w}val', 'true')
            row_bg = HEADER_BG if is_header else (ALT_ROW_BG if row_index % 2 == 1 else None)
            for col_idx in range(num_cols):
                tc = ET.SubElement(tr, f'{w}tc')
                tcPr = ET.SubElement(tc, f'{w}tcPr')
                tcW_el = ET.SubElement(tcPr, f'{w}tcW')
                tcW_el.set(f'{w}w', str(col_widths[col_idx]))
                tcW_el.set(f'{w}type', 'dxa')
                if row_bg:
                    shd = ET.SubElement(tcPr, f'{w}shd')
                    shd.set(f'{w}val',   'clear')
                    shd.set(f'{w}color', 'auto')
                    shd.set(f'{w}fill',  row_bg)
                tcMar = ET.SubElement(tcPr, f'{w}tcMar')
                for side, pts in (('top','60'),('bottom','60'),('left','108'),('right','108')):
                    m = ET.SubElement(tcMar, f'{w}{side}')
                    m.set(f'{w}w',    pts)
                    m.set(f'{w}type', 'dxa')
                p = ET.SubElement(tc, f'{w}p')
                pPr = ET.SubElement(p, f'{w}pPr')
                pStyle = ET.SubElement(pPr, f'{w}pStyle')
                pStyle.set(f'{w}val', 'Normal')
                spacing = ET.SubElement(pPr, f'{w}spacing')
                spacing.set(f'{w}before', '40')
                spacing.set(f'{w}after',  '40')
                cell_text = cells[col_idx] if col_idx < len(cells) else ''
                self._create_table_cell_runs(cell_text, p, is_header=is_header)

        if header:
            add_row(header, is_header=True)
        for idx, row in enumerate(data_rows):
            add_row(row, is_header=False, row_index=idx)

        return tbl

    def _create_spacer(self) -> ET.Element:
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        p = ET.Element(f'{w}p')
        pPr = ET.SubElement(p, f'{w}pPr')
        sp = ET.SubElement(pPr, f'{w}spacing')
        sp.set(f'{w}before', '80')
        sp.set(f'{w}after',  '80')
        return p

    # ------------------------------------------------------------------ #
    # Cover page                                                           #
    # ------------------------------------------------------------------ #

    def _build_cover_elements(self, template_dir: str, title: str) -> list:
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

        # Full-width blue table, single tall cell
        tbl = ET.Element(f'{w}tbl')
        tblPr = ET.SubElement(tbl, f'{w}tblPr')
        tblW = ET.SubElement(tblPr, f'{w}tblW')
        tblW.set(f'{w}w', '5000')
        tblW.set(f'{w}type', 'pct')
        tblBorders = ET.SubElement(tblPr, f'{w}tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = ET.SubElement(tblBorders, f'{w}{side}')
            b.set(f'{w}val',   'none')
            b.set(f'{w}sz',    '0')
            b.set(f'{w}space', '0')
            b.set(f'{w}color', 'auto')
        tblGrid = ET.SubElement(tbl, f'{w}tblGrid')
        ET.SubElement(tblGrid, f'{w}gridCol')

        tr = ET.SubElement(tbl, f'{w}tr')
        trPr = ET.SubElement(tr, f'{w}trPr')
        trH = ET.SubElement(trPr, f'{w}trHeight')
        trH.set(f'{w}val',   '12850')   # ~22.5cm, fills A4 content area
        trH.set(f'{w}hRule', 'exact')

        tc = ET.SubElement(tr, f'{w}tc')
        tcPr = ET.SubElement(tc, f'{w}tcPr')
        tcW = ET.SubElement(tcPr, f'{w}tcW')
        tcW.set(f'{w}w', '5000')
        tcW.set(f'{w}type', 'pct')
        shd = ET.SubElement(tcPr, f'{w}shd')
        shd.set(f'{w}val',   'clear')
        shd.set(f'{w}color', 'auto')
        shd.set(f'{w}fill',  '3313E2')
        vAlign = ET.SubElement(tcPr, f'{w}vAlign')
        vAlign.set(f'{w}val', 'center')
        tcMar = ET.SubElement(tcPr, f'{w}tcMar')
        for side, val in (('top','1440'),('left','720'),('bottom','1440'),('right','720')):
            m = ET.SubElement(tcMar, f'{w}{side}')
            m.set(f'{w}w',    val)
            m.set(f'{w}type', 'dxa')

        # --- Logo ---
        logo_path = str(LOGO_PATH)
        if os.path.exists(logo_path):
            try:
                logo_rid = self._embed_image(template_dir, logo_path, 'cover_logo')
                try:
                    img_w, img_h = self._read_png_dimensions(logo_path)
                except Exception:
                    img_w, img_h = 200, 80
                target_cx = 1260000  # 3.5 cm in EMU
                target_cy = int(img_h * target_cx / img_w) if img_w > 0 else 504000
                logo_p = ET.SubElement(tc, f'{w}p')
                logo_pPr = ET.SubElement(logo_p, f'{w}pPr')
                jc = ET.SubElement(logo_pPr, f'{w}jc')
                jc.set(f'{w}val', 'center')
                sp = ET.SubElement(logo_pPr, f'{w}spacing')
                sp.set(f'{w}before', '0')
                sp.set(f'{w}after',  '360')
                self._add_image_run(logo_p, logo_rid, target_cx, target_cy)
            except Exception:
                self._cover_text(tc, 'XD.AI', 48, bold=True)

        # --- Title ---
        self._cover_text(tc, title, 72, bold=True)

        # --- Divider ---
        div_p = ET.SubElement(tc, f'{w}p')
        div_pPr = ET.SubElement(div_p, f'{w}pPr')
        div_sp = ET.SubElement(div_pPr, f'{w}spacing')
        div_sp.set(f'{w}before', '200')
        div_sp.set(f'{w}after',  '200')
        pBdr = ET.SubElement(div_pPr, f'{w}pBdr')
        bot = ET.SubElement(pBdr, f'{w}bottom')
        bot.set(f'{w}val',   'single')
        bot.set(f'{w}sz',    '6')
        bot.set(f'{w}space', '1')
        bot.set(f'{w}color', 'FFFFFF')

        # --- Date ---
        date_str = datetime.date.today().strftime('%B %Y')
        self._cover_text(tc, date_str, 24, bold=False)

        # Page break after cover
        pb_p = ET.Element(f'{w}p')
        pb_pPr = ET.SubElement(pb_p, f'{w}pPr')
        pb_sp = ET.SubElement(pb_pPr, f'{w}spacing')
        pb_sp.set(f'{w}before', '0')
        pb_sp.set(f'{w}after',  '0')
        pb_r = ET.SubElement(pb_p, f'{w}r')
        pb_br = ET.SubElement(pb_r, f'{w}br')
        pb_br.set(f'{w}type', 'page')

        return [tbl, pb_p]

    def _cover_text(self, parent, text: str, sz_half: int, bold=False):
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        p = ET.SubElement(parent, f'{w}p')
        pPr = ET.SubElement(p, f'{w}pPr')
        jc = ET.SubElement(pPr, f'{w}jc')
        jc.set(f'{w}val', 'center')
        sp = ET.SubElement(pPr, f'{w}spacing')
        sp.set(f'{w}before', '80')
        sp.set(f'{w}after',  '80')
        run = ET.SubElement(p, f'{w}r')
        rPr = ET.SubElement(run, f'{w}rPr')
        rFonts = ET.SubElement(rPr, f'{w}rFonts')
        rFonts.set(f'{w}ascii', 'Causten')
        rFonts.set(f'{w}hAnsi', 'Causten')
        if bold:
            ET.SubElement(rPr, f'{w}b')
        el_sz = ET.SubElement(rPr, f'{w}sz')
        el_sz.set(f'{w}val', str(sz_half))
        el_szCs = ET.SubElement(rPr, f'{w}szCs')
        el_szCs.set(f'{w}val', str(sz_half))
        clr = ET.SubElement(rPr, f'{w}color')
        clr.set(f'{w}val', 'FFFFFF')
        t = ET.SubElement(run, f'{w}t')
        t.text = text

    # ------------------------------------------------------------------ #
    # Table of Contents                                                    #
    # ------------------------------------------------------------------ #

    def _create_toc_elements(self) -> list:
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        elements = []

        # "Contents" heading
        h = ET.Element(f'{w}p')
        hPr = ET.SubElement(h, f'{w}pPr')
        hStyle = ET.SubElement(hPr, f'{w}pStyle')
        hStyle.set(f'{w}val', 'Heading1')
        hRun = ET.SubElement(h, f'{w}r')
        hT = ET.SubElement(hRun, f'{w}t')
        hT.text = 'Contents'
        elements.append(h)

        # TOC field
        toc_p = ET.Element(f'{w}p')
        r1 = ET.SubElement(toc_p, f'{w}r')
        fc1 = ET.SubElement(r1, f'{w}fldChar')
        fc1.set(f'{w}fldCharType', 'begin')
        fc1.set(f'{w}dirty',       'true')
        r2 = ET.SubElement(toc_p, f'{w}r')
        instr = ET.SubElement(r2, f'{w}instrText')
        instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instr.text = ' TOC \\o "1-3" \\h \\z \\u '
        r3 = ET.SubElement(toc_p, f'{w}r')
        fc2 = ET.SubElement(r3, f'{w}fldChar')
        fc2.set(f'{w}fldCharType', 'separate')
        r4 = ET.SubElement(toc_p, f'{w}r')
        ph = ET.SubElement(r4, f'{w}t')
        ph.text = 'Right-click this text in Word and choose Update Field to generate the table of contents.'
        r5 = ET.SubElement(toc_p, f'{w}r')
        fc3 = ET.SubElement(r5, f'{w}fldChar')
        fc3.set(f'{w}fldCharType', 'end')
        elements.append(toc_p)

        # Page break after TOC
        pb = ET.Element(f'{w}p')
        pb_r = ET.SubElement(pb, f'{w}r')
        pb_br = ET.SubElement(pb_r, f'{w}br')
        pb_br.set(f'{w}type', 'page')
        elements.append(pb)

        return elements

    # ------------------------------------------------------------------ #
    # Image embedding                                                      #
    # ------------------------------------------------------------------ #

    def _embed_image(self, template_dir: str, image_path: str, image_name: str) -> str:
        word_dir  = os.path.join(template_dir, 'word')
        media_dir = os.path.join(word_dir, 'media')
        os.makedirs(media_dir, exist_ok=True)

        ext       = os.path.splitext(image_path)[1].lower()
        dest_name = f'{image_name}{ext}'
        shutil.copy2(image_path, os.path.join(media_dir, dest_name))

        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_content = f.read()
        rids     = re.findall(r'Id="rId(\d+)"', rels_content)
        next_rid = max(int(r) for r in rids) + 1 if rids else 10
        image_rid = f'rId{next_rid}'
        img_rel = (
            f'<Relationship Id="{image_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{dest_name}"/>'
        )
        rels_content = rels_content.replace('</Relationships>', f'{img_rel}</Relationships>')
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)

        ct_path = os.path.join(template_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct_content = f.read()
        ext_key = ext.lstrip('.')
        mime    = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg'}.get(ext_key, 'image/png')
        if f'Extension="{ext_key}"' not in ct_content:
            ct_content = ct_content.replace('</Types>', f'<Default Extension="{ext_key}" ContentType="{mime}"/></Types>')
            with open(ct_path, 'w', encoding='utf-8') as f:
                f.write(ct_content)

        return image_rid

    def _add_image_run(self, parent: ET.Element, rid: str, cx: int, cy: int, doc_pr_id: int = 2001):
        w      = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        wp     = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
        a      = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
        pic_ns = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
        r_ns   = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

        run    = ET.SubElement(parent, f'{w}r')
        drawing = ET.SubElement(run, f'{w}drawing')
        inline  = ET.SubElement(drawing, f'{wp}inline')
        for attr in ('distT','distB','distL','distR'):
            inline.set(attr, '0')

        extent = ET.SubElement(inline, f'{wp}extent')
        extent.set('cx', str(cx))
        extent.set('cy', str(cy))
        eff = ET.SubElement(inline, f'{wp}effectExtent')
        for attr in ('l','t','r','b'):
            eff.set(attr, '0')
        docPr = ET.SubElement(inline, f'{wp}docPr')
        docPr.set('id',   str(doc_pr_id))
        docPr.set('name', 'Logo')
        ET.SubElement(inline, f'{wp}cNvGraphicFramePr')

        graphic     = ET.SubElement(inline, f'{a}graphic')
        graphicData = ET.SubElement(graphic, f'{a}graphicData')
        graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

        pic_el  = ET.SubElement(graphicData, f'{pic_ns}pic')
        nvPicPr = ET.SubElement(pic_el, f'{pic_ns}nvPicPr')
        cNvPr   = ET.SubElement(nvPicPr, f'{pic_ns}cNvPr')
        cNvPr.set('id',   str(doc_pr_id + 1))
        cNvPr.set('name', 'Logo')
        ET.SubElement(nvPicPr, f'{pic_ns}cNvPicPr')

        blipFill = ET.SubElement(pic_el, f'{pic_ns}blipFill')
        blip     = ET.SubElement(blipFill, f'{a}blip')
        blip.set(f'{r_ns}embed', rid)
        stretch  = ET.SubElement(blipFill, f'{a}stretch')
        ET.SubElement(stretch, f'{a}fillRect')

        spPr = ET.SubElement(pic_el, f'{pic_ns}spPr')
        xfrm = ET.SubElement(spPr, f'{a}xfrm')
        off  = ET.SubElement(xfrm, f'{a}off')
        off.set('x', '0')
        off.set('y', '0')
        ext = ET.SubElement(xfrm, f'{a}ext')
        ext.set('cx', str(cx))
        ext.set('cy', str(cy))
        prstGeom = ET.SubElement(spPr, f'{a}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{a}avLst')

    @staticmethod
    def _read_png_dimensions(path: str):
        with open(path, 'rb') as f:
            f.read(16)  # PNG sig (8) + IHDR length (4) + 'IHDR' (4)
            width  = struct.unpack('>I', f.read(4))[0]
            height = struct.unpack('>I', f.read(4))[0]
        return width, height

    # ------------------------------------------------------------------ #
    # Document assembly                                                    #
    # ------------------------------------------------------------------ #

    def _extract_title(self, paragraphs: list) -> str:
        for p in paragraphs:
            if p.get('type') == 'heading' and p.get('level') == 1:
                return p.get('content', 'Document')
        return 'Document'

    def _update_document(self, template_dir: str, paragraphs: list):
        doc_path = os.path.join(template_dir, 'word', 'document.xml')

        with open(doc_path, 'r', encoding='utf-8') as f:
            original_content = f.read()

        doc_tag_match  = re.search(r'<w:document[^>]+>', original_content)
        original_doc_tag = doc_tag_match.group(0) if doc_tag_match else None

        tree = ET.parse(doc_path)
        root = tree.getroot()
        w    = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

        body = root.find(f'.//{w}body')
        if body is None:
            raise ValueError("Could not find document body")

        sectPr = body.find(f'{w}sectPr')
        for child in list(body):
            if child.tag != f'{w}sectPr':
                body.remove(child)

        # Build ordered list of all elements
        all_elements = []

        if self.cover:
            title = self._extract_title(paragraphs)
            all_elements.extend(self._build_cover_elements(template_dir, title))

        if self.toc:
            all_elements.extend(self._create_toc_elements())

        for para_data in paragraphs:
            if para_data['type'] == 'table':
                all_elements.append(self._create_table(para_data))
                all_elements.append(self._create_spacer())
            else:
                all_elements.append(self._create_paragraph(para_data))

        # Insert before sectPr
        if sectPr is not None:
            base = list(body).index(sectPr)
            for i, element in enumerate(all_elements):
                body.insert(base + i, element)
        else:
            for element in all_elements:
                body.append(element)

        tree.write(doc_path, xml_declaration=True, encoding='UTF-8')

        if original_doc_tag:
            with open(doc_path, 'r', encoding='utf-8') as f:
                new_content = f.read()
            new_doc_tag_match = re.search(r'<w:document[^>]+>', new_content)
            if new_doc_tag_match:
                new_tag = new_doc_tag_match.group(0)
                # Merge any namespace declarations added by ET (e.g. xmlns:a, xmlns:pic for images)
                # that are not present in the original template tag
                new_ns = dict(re.findall(r'xmlns:(\w+)="([^"]*)"', new_tag))
                orig_ns = dict(re.findall(r'xmlns:(\w+)="([^"]*)"', original_doc_tag))
                missing = {k: v for k, v in new_ns.items() if k not in orig_ns}
                if missing:
                    extra = ' '.join(f'xmlns:{k}="{v}"' for k, v in missing.items())
                    merged_tag = original_doc_tag[:-1] + ' ' + extra + '>'
                else:
                    merged_tag = original_doc_tag
                new_content = new_content.replace(new_tag, merged_tag, 1)
                with open(doc_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)

    # ------------------------------------------------------------------ #
    # Footer                                                               #
    # ------------------------------------------------------------------ #

    def _add_footer(self, template_dir: str):
        word_dir = os.path.join(template_dir, 'word')

        footer_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Footer"/>
            <w:jc w:val="right"/>
            <w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
        </w:pPr>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">Page </w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>1</w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>1</w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Causten" w:hAnsi="Causten"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
</w:ftr>'''

        footer_path = os.path.join(word_dir, 'footer1.xml')
        with open(footer_path, 'w', encoding='utf-8') as f:
            f.write(footer_xml)

        rels_path = os.path.join(word_dir, '_rels', 'document.xml.rels')
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_content = f.read()
        rids     = re.findall(r'Id="rId(\d+)"', rels_content)
        next_rid = max(int(r) for r in rids) + 1 if rids else 10
        footer_rid = f'rId{next_rid}'
        footer_rel = (
            f'<Relationship Id="{footer_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
            f'Target="footer1.xml"/>'
        )
        rels_content = rels_content.replace('</Relationships>', f'{footer_rel}</Relationships>')
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)

        doc_path = os.path.join(word_dir, 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc_content = f.read()
        ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        footer_ref = (
            f'<w:footerReference xmlns:w="{ns_w}" xmlns:r="{ns_r}" '
            f'w:type="default" r:id="{footer_rid}"/>'
        )
        if '<w:headerReference' in doc_content:
            doc_content = re.sub(r'(<w:headerReference[^/]*/\s*>)', r'\1' + footer_ref, doc_content)
        elif '<w:sectPr' in doc_content:
            doc_content = re.sub(r'(<w:sectPr[^>]*>)', r'\1' + footer_ref, doc_content)
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_content)

        ct_path = os.path.join(template_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct_content = f.read()
        if 'footer1.xml' not in ct_content:
            ct_content = ct_content.replace(
                '</Types>',
                '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'
            )
            with open(ct_path, 'w', encoding='utf-8') as f:
                f.write(ct_content)

    # ------------------------------------------------------------------ #
    # DOCX packaging                                                       #
    # ------------------------------------------------------------------ #

    def _create_docx(self, template_dir: str, output_path: str):
        ct_path = os.path.join(template_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct_content = f.read()
        ct_content = ct_content.replace(
            'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
        )
        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct_content)

        if os.path.exists(output_path):
            os.remove(output_path)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, _, files in os.walk(template_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname   = os.path.relpath(file_path, template_dir)
                    zf.write(file_path, arcname)


# ---------------------------------------------------------------------- #
# CLI                                                                     #
# ---------------------------------------------------------------------- #

def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to branded DOCX (XD.AI)')
    parser.add_argument('input',       help='Input Markdown file')
    parser.add_argument('output',      help='Output DOCX file')
    parser.add_argument('--template',  help='Custom template path', default=None)
    parser.add_argument('--cover',     action='store_true', help='Add a branded cover page')
    parser.add_argument('--toc',       action='store_true', help='Add a table of contents')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        md_content = f.read()

    converter = MarkdownToDocx(
        template_path=args.template,
        cover=args.cover,
        toc=args.toc,
    )
    output = converter.convert(md_content, args.output)
    print(f"Created: {output}")


if __name__ == '__main__':
    main()
