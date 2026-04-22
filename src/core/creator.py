import os
import re
import marko
from pathlib import Path
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# Библиотеки для подсветки кода
try:
    from pygments import lexers
    from pygments.lexers import get_lexer_by_name
    from pygments.token import Token
except ImportError:
    Token = None

from .errors import PptxSyntaxError

# Загрузка .env
env_path = Path(__file__).resolve().parent.parent.parent / '.env'
load_dotenv(dotenv_path=env_path)

# ЯРКАЯ ПАЛИТРА ДЛЯ БЕЛОГО ФОНА
CODE_THEME = {
    'Keyword': RGBColor(0, 0, 255),
    'Name': RGBColor(0, 0, 0),
    'String': RGBColor(163, 21, 21),
    'Comment': RGBColor(0, 128, 0),
    'Operator': RGBColor(0, 0, 0),
    'Number': RGBColor(9, 134, 88),
    'Type': RGBColor(38, 127, 153),
    'Other': RGBColor(0, 0, 0)
}


class PptxCreator:
    def __init__(self):
        self.prs = Presentation()

        # Основные шрифты (НЕ МЕНЯЕМ)
        self.font_name = os.getenv("PPT_FONT", "Bookman Old Style")
        self.title_size = int(os.getenv("PPT_TITLE_SIZE", "30"))
        self.body_size = int(os.getenv("PPT_BODY_SIZE", "22"))

        # НОВЫЕ НАСТРОЙКИ ДЛЯ КОДА ИЗ ENV
        self.code_font = os.getenv("PPT_CODE_FONT", "Courier New")
        self.code_size = int(os.getenv("PPT_CODE_SIZE", "20"))

        self._set_aspect_ratio(os.getenv("PPT_ASPECT_RATIO", "16:9"))
        self.title_layout = self.prs.slide_layouts[0]
        self.content_layout = self.prs.slide_layouts[1]
        self.warnings = []

    def _set_aspect_ratio(self, ratio_str):
        ratios = {"4:3": (9144000, 6858000), "16:9": (12192000, 6858000)}
        w, h = ratios.get(ratio_str, ratios["16:9"])
        self.prs.slide_width, self.prs.slide_height = w, h

    def _remove_bullet(self, paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        from pptx.oxml.xmlchemy import OxmlElement
        buNone = OxmlElement('a:buNone')
        pPr.insert(0, buNone)

    def _setup_text_frame(self, shape, align=None, is_title=False):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if is_title else MSO_ANCHOR.TOP
        tf.clear()
        p = tf.paragraphs[0]
        if align: p.alignment = align
        return tf

    def _position_shape(self, shape, top_inch, height_inch, width_percent=0.9):
        slide_width = self.prs.slide_width
        new_width = int(slide_width * width_percent)
        shape.width, shape.height = new_width, Inches(height_inch)
        shape.left, shape.top = int((slide_width - new_width) / 2), Inches(top_inch)
        return shape

    def _add_highlighted_code(self, text_frame, code_text, lang='text'):
        """Раскрашивает код, используя настройки из ENV."""
        try:
            lexer = get_lexer_by_name(lang if lang else 'text', stripall=True)
        except:
            lexer = lexers.get_lexer_by_name('text')

        lines = code_text.split('\n')
        for line in lines:
            if not line.strip() and line == lines[-1]: continue

            p = text_frame.add_paragraph()
            self._remove_bullet(p)

            tokens = lexer.get_tokens(line)
            for ttype, value in tokens:
                clean_value = value.replace('\r', '').replace('\n', '')
                if not clean_value and value: continue

                run = p.add_run()
                run.text = clean_value
                run.font.name = self.code_font  # ИЗ ENV
                run.font.size = Pt(self.code_size)  # ИЗ ENV

                if ttype in Token.Keyword:
                    color = CODE_THEME['Keyword']
                elif ttype in Token.Name:
                    color = CODE_THEME['Name']
                elif ttype in Token.Literal.String:
                    color = CODE_THEME['String']
                elif ttype in Token.Comment:
                    color = CODE_THEME['Comment']
                elif ttype in Token.Operator:
                    color = CODE_THEME['Operator']
                elif ttype in Token.Literal.Number:
                    color = CODE_THEME['Number']
                elif ttype in Token.Name.Type or ttype in Token.Name.Class:
                    color = CODE_THEME['Type']
                else:
                    color = CODE_THEME['Other']

                run.font.color.rgb = color

    def _add_node_to_frame(self, text_frame, node, level=0, default_align=None):
        ntype = node.__class__.__name__

        if ntype == 'Paragraph':
            p = text_frame.add_paragraph()
            p.level = min(level, 8)
            if default_align: p.alignment = default_align
            self._fill_run(p, node)

        elif ntype == 'List':
            for item in node.children:
                if item.__class__.__name__ == 'ListItem':
                    for sub in item.children:
                        if sub.__class__.__name__ == 'List':
                            self._add_node_to_frame(text_frame, sub, level=level + 1)
                        else:
                            self._add_node_to_frame(text_frame, sub, level=level)

        elif ntype in ['FencedCode', 'CodeBlock']:
            lang = getattr(node, 'lang', 'text')
            content = node.children[0].children if hasattr(node.children[0], 'children') else node.children[0]
            self._add_highlighted_code(text_frame, str(content), lang)

    def _create_title_slide(self, doc):
        slide = self.prs.slides.add_slide(self.title_layout)
        title_node, other_nodes = None, []
        for node in doc.children:
            if node.__class__.__name__ == 'Heading':
                title_node = node
            else:
                other_nodes.append(node)

        if slide.shapes.title:
            shape = self._position_shape(slide.shapes.title, top_inch=2.0, height_inch=2.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            if title_node: self._fill_run(tf.paragraphs[0], title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=4.2, height_inch=2.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER)
            for node in other_nodes: self._add_node_to_frame(tf, node, default_align=PP_ALIGN.CENTER)

    def _create_content_slide(self, doc, slide_num):
        slide = self.prs.slides.add_slide(self.content_layout)
        title_node, content_nodes = None, []
        for node in doc.children:
            if node.__class__.__name__ == 'Heading' and not title_node:
                title_node = node
            else:
                content_nodes.append(node)

        if slide.shapes.title:
            shape = self._position_shape(slide.shapes.title, top_inch=0.4, height_inch=1.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            if title_node: self._fill_run(tf.paragraphs[0], title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=1.6, height_inch=5.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.LEFT)
            for node in content_nodes: self._add_node_to_frame(tf, node)

    def _fill_run(self, paragraph, node, is_title=False):
        if not hasattr(node, 'children'): return
        for child in node.children:
            ctype = child.__class__.__name__
            if isinstance(child, str):
                self._create_run(paragraph, child, is_title)
            elif ctype == 'RawText':
                self._create_run(paragraph, child.children, is_title)
            elif ctype == 'Strong':
                self._apply_style(paragraph, child, is_title, bold=True)
            elif ctype == 'Emphasis':
                self._apply_style(paragraph, child, is_title, italic=True)
            elif ctype == 'CodeSpan':
                self._create_run(paragraph, child.children, is_title, code=True)
            elif hasattr(child, 'children'):
                self._fill_run(paragraph, child, is_title)

    def _apply_style(self, paragraph, node, is_title, bold=False, italic=False):
        text = node.children[0] if isinstance(node.children, list) and isinstance(node.children[0],
                                                                                  str) else node.children
        if isinstance(text, str):
            self._create_run(paragraph, text, is_title, bold=bold, italic=italic)
        elif hasattr(node, 'children'):
            for sub in node.children:
                content = getattr(sub, 'children', str(sub))
                self._create_run(paragraph, content, is_title, bold=bold, italic=italic)

    def _create_run(self, paragraph, text, is_title, bold=False, italic=False, code=False):
        if not text or not str(text).strip(): return
        run = paragraph.add_run()
        run.text = str(text)

        # ПРИМЕНЕНИЕ ШРИФТОВ
        if code:
            run.font.name = self.code_font
            run.font.size = Pt(self.code_size)
        else:
            run.font.name = self.font_name
            run.font.size = Pt(self.title_size if is_title else self.body_size)

        run.font.bold, run.font.italic = bold, italic

    def create_from_text(self, md_text, output_path):
        self.warnings = []
        blocks = re.split(r'\n\s*---\s*\n', md_text.strip())
        if not blocks or not blocks[0].strip(): raise PptxSyntaxError("MD текст пуст.")
        for idx, block in enumerate(blocks):
            if not block.strip(): continue
            doc = marko.parse(block)
            if idx == 0:
                self._create_title_slide(doc)
            else:
                self._create_content_slide(doc, idx + 1)
        self.prs.save(output_path)
        return {"slides_created": len(self.prs.slides), "warnings": self.warnings}

    def create_from_file(self, md_path, output_path):
        with open(md_path, 'r', encoding='utf-8') as f: return self.create_from_text(f.read(), output_path)