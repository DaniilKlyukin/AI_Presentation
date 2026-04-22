import os
import re
import uuid
import marko
from pathlib import Path
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

try:
    import matplotlib.pyplot as plt
    from matplotlib import rcParams

    rcParams['mathtext.fontset'] = 'cm'
    plt.switch_backend('Agg')
except ImportError:
    plt = None

try:
    from pygments import lexers
    from pygments.lexers import get_lexer_by_name
    from pygments.token import Token
except ImportError:
    Token = None

from .errors import PptxSyntaxError

env_path = Path(__file__).resolve().parent.parent.parent / '.env'
load_dotenv(dotenv_path=env_path)

CODE_THEME = {
    'Keyword': RGBColor(0, 0, 255),
    'Name': RGBColor(0, 0, 0),
    'Name.Function': RGBColor(121, 94, 38),
    'Name.Class': RGBColor(38, 127, 153),
    'String': RGBColor(163, 21, 21),
    'Comment': RGBColor(0, 128, 0),
    'Operator': RGBColor(0, 0, 0),
    'Number': RGBColor(9, 134, 88),
    'Type': RGBColor(38, 127, 153),
    'Other': RGBColor(0, 0, 0)
}


# Вспомогательный класс для обработки частичных узлов текста
class DummyNode:
    def __init__(self, children):
        self.children = children


class PptxCreator:
    def __init__(self):
        self.prs = Presentation()
        self.font_name = os.getenv("PPT_FONT", "Bookman Old Style")
        self.title_size = int(os.getenv("PPT_TITLE_SIZE", "30"))
        self.body_size = int(os.getenv("PPT_BODY_SIZE", "22"))
        self.code_font = os.getenv("PPT_CODE_FONT", "Courier New")
        self.code_size = int(os.getenv("PPT_CODE_SIZE", "20"))
        self.line_spacing = float(os.getenv("PPT_LINE_SPACING", "1.0"))
        self._set_aspect_ratio(os.getenv("PPT_ASPECT_RATIO", "16:9"))
        self.title_layout = self.prs.slide_layouts[0]
        self.content_layout = self.prs.slide_layouts[1]
        self.warnings = []
        self.temp_files = []
        self.md_parser = marko.Markdown(extensions=['gfm'])

    def _set_aspect_ratio(self, ratio_str):
        ratios = {"4:3": (9144000, 6858000), "16:9": (12192000, 6858000)}
        w, h = ratios.get(ratio_str, ratios["16:9"])
        self.prs.slide_width, self.prs.slide_height = w, h

    # ---------------------- ФОРМУЛЫ -------------------------
    def _render_formula_to_image(self, formula_text):
        if not plt:
            self.warnings.append("Установите matplotlib для рендера формул")
            return None
        formula = formula_text.strip().replace('$', '')
        filename = f"math_{uuid.uuid4().hex}.png"
        fig = plt.figure(figsize=(8, 1), dpi=300)
        try:
            fig.text(0.5, 0.5, f"${formula}$", fontsize=self.body_size + 4,
                     ha='center', va='center')
            # Делаем pad_inches минимальным (0.01), чтобы избежать пустых полей
            fig.savefig(filename, transparent=True, bbox_inches='tight', pad_inches=0.01)
            plt.close(fig)
            self.temp_files.append(filename)
            return filename
        except Exception as e:
            self.warnings.append(f"Math error: {e}")
            plt.close(fig)
            return None

    def _process_math_blocks(self, text):
        def replacer(m):
            img_path = self._render_formula_to_image(m.group(1))
            return f' ![math]({img_path}) ' if img_path else m.group(0)

        # Сначала блочные (оборачиваем в \n для надежности)
        text = re.sub(r'\$\$(.*?)\$\$', lambda m: f'\n{replacer(m)}\n', text, flags=re.DOTALL)
        # Затем инлайн
        text = re.sub(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', replacer, text)
        return text

    # ---------------------- ВСПОМОГАТЕЛЬНЫЕ -----------------
    def _apply_paragraph_style(self, paragraph, is_code=False, align=None):
        if align:
            paragraph.alignment = align
        paragraph.line_spacing = self.line_spacing
        if is_code:
            self._remove_bullet_xml(paragraph)

    def _remove_bullet_xml(self, paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        pPr.insert(0, OxmlElement('a:buNone'))
        paragraph.left_indent = 0
        paragraph.first_line_indent = 0

    def _setup_text_frame(self, shape, align=None, is_title=False):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if is_title else MSO_ANCHOR.TOP
        tf.clear()
        return tf

    def _position_shape(self, shape, top_inch, height_inch=None, width_percent=0.9):
        slide_width = self.prs.slide_width
        new_width = int(slide_width * width_percent)
        shape.width = new_width
        if height_inch:
            shape.height = Inches(height_inch)
        shape.left = int((slide_width - new_width) / 2)
        shape.top = Inches(top_inch)
        return shape

    def _get_or_add_paragraph(self, text_frame):
        if len(text_frame.paragraphs) == 1 and not text_frame.paragraphs[0].text:
            return text_frame.paragraphs[0]
        return text_frame.add_paragraph()

    def _get_current_tf_height(self, text_frame):
        """Максимально точно вычисляет текущую высоту текста во фрейме в дюймах"""
        total_pt = 0
        for p in text_frame.paragraphs:
            if p.space_before:
                total_pt += p.space_before.pt

            text = p.text.strip()
            max_size = self.body_size
            for run in p.runs:
                if run.font.size:
                    max_size = max(max_size, run.font.size.pt)

            if not text:
                total_pt += max_size
            else:
                # Оценка длины строки (около 65 символов для размера 22 на слайде 16:9)
                chars_per_line = max(40, int(1400 / max_size))
                lines = (len(text) // chars_per_line) + 1
                total_pt += lines * max_size * self.line_spacing * 1.15  # 1.15 - интерлиньяж pptx

            if p.space_after:
                total_pt += p.space_after.pt

        return total_pt / 72.0

    # ---------------------- ВСТАВКА ИЗОБРАЖЕНИЙ И ТАБЛИЦ -----
    def _insert_image_shape(self, slide, text_frame, img_path):
        if not os.path.exists(img_path):
            return

        current_text_h = self._get_current_tf_height(text_frame)
        top_inch = text_frame._parent.top.inches + current_text_h + 0.05  # 0.05 дюйма отступ

        pic = slide.shapes.add_picture(img_path, Inches(0), Inches(top_inch))

        max_width = self.prs.slide_width * 0.9
        if pic.width > max_width:
            ratio = max_width / pic.width
            pic.width = int(max_width)
            pic.height = int(pic.height * ratio)

        pic.left = int((self.prs.slide_width - pic.width) / 2)

        # Создаем прозрачный микро-абзац, отступ которого равен высоте картинки!
        p = text_frame.add_paragraph()
        self._remove_bullet_xml(p)
        p.text = " "
        p.font.size = Pt(1)
        # Отступ равен высоте картинки + 10 пунктов для эстетичного зазора
        p.space_before = Pt(pic.height.pt + 10)

    def _add_table(self, slide, text_frame, node):
        rows_nodes = node.children
        rows = len(rows_nodes)
        cols = max((len(r.children) for r in rows_nodes), default=0)
        if rows == 0 or cols == 0:
            return

        current_text_h = self._get_current_tf_height(text_frame)
        top_inch = text_frame._parent.top.inches + current_text_h + 0.1

        width = int(self.prs.slide_width * 0.9)
        left = int((self.prs.slide_width - width) / 2)
        row_height = 0.4
        tbl_shape = slide.shapes.add_table(rows, cols, left, Inches(top_inch), width, Inches(row_height * rows))
        table = tbl_shape.table

        for r_idx, row_node in enumerate(rows_nodes):
            for c_idx, cell_node in enumerate(row_node.children):
                if c_idx < cols:
                    cell_tf = table.cell(r_idx, c_idx).text_frame
                    cell_tf.clear()
                    p = cell_tf.paragraphs[0]
                    self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
                    self._fill_run(p, cell_node)
                    for run in p.runs:
                        run.font.name = self.font_name
                        run.font.size = Pt(self.body_size - 4)
                        if r_idx == 0:
                            run.font.bold = True

        # Сдвиг текста под таблицу
        p = text_frame.add_paragraph()
        self._remove_bullet_xml(p)
        p.text = " "
        p.font.size = Pt(1)
        p.space_before = Pt(tbl_shape.height.pt + 15)

    # ---------------------- ПОДСВЕТКА КОДА -------------------
    def _add_highlighted_code(self, text_frame, code_text, lang='text'):
        code_text = code_text.replace('\t', '    ').strip('\n\r')
        try:
            lexer = get_lexer_by_name(lang)
        except:
            lexer = lexers.get_lexer_by_name('text')

        lines = code_text.split('\n')
        for i, line in enumerate(lines):
            p = self._get_or_add_paragraph(text_frame) if i == 0 else text_frame.add_paragraph()
            self._apply_paragraph_style(p, is_code=True, align=PP_ALIGN.LEFT)
            tokens = lexer.get_tokens(line)
            for ttype, value in tokens:
                val = value.replace('\r', '').replace('\n', '')
                if not val:
                    continue
                run = p.add_run()
                run.text = val
                run.font.name = self.code_font
                run.font.size = Pt(self.code_size)
                if ttype in Token.Keyword:
                    color = CODE_THEME['Keyword']
                elif ttype in Token.Name.Function:
                    color = CODE_THEME['Name.Function']
                elif ttype in Token.Name.Class:
                    color = CODE_THEME['Name.Class']
                elif ttype in Token.Literal.String:
                    color = CODE_THEME['String']
                elif ttype in Token.Comment:
                    color = CODE_THEME['Comment']
                elif ttype in Token.Literal.Number:
                    color = CODE_THEME['Number']
                elif ttype in Token.Operator:
                    color = CODE_THEME['Operator']
                elif ttype in Token.Name.Type:
                    color = CODE_THEME['Type']
                else:
                    color = CODE_THEME['Other']
                run.font.color.rgb = color

    # ---------------------- ОБХОД AST ------------------------
    def _add_node_to_frame(self, text_frame, node, slide=None, level=0, default_align=None):
        ntype = node.__class__.__name__

        if ntype == 'Paragraph':
            p = self._get_or_add_paragraph(text_frame)
            p.level = min(level, 8)
            self._apply_paragraph_style(p, align=default_align)

            for child in getattr(node, 'children', []):
                # Если встретили формулу/картинку (инлайн или блочную)
                if child.__class__.__name__ == 'Image':
                    if slide:
                        self._insert_image_shape(slide, text_frame, child.dest)

                    # Начинаем новую строку для текста ПОСЛЕ формулы, НО отключаем маркер списка
                    p = text_frame.add_paragraph()
                    p.level = min(level, 8)
                    self._remove_bullet_xml(p)
                    self._apply_paragraph_style(p, align=default_align)
                else:
                    self._fill_run(p, DummyNode([child]))

        elif ntype == 'List':
            for item in node.children:
                if item.__class__.__name__ == 'ListItem':
                    for sub in item.children:
                        if sub.__class__.__name__ == 'List':
                            self._add_node_to_frame(text_frame, sub, slide=slide, level=level + 1)
                        else:
                            self._add_node_to_frame(text_frame, sub, slide=slide, level=level,
                                                    default_align=default_align)

        elif ntype in ['FencedCode', 'CodeBlock']:
            lang = getattr(node, 'lang', 'text')
            content = node.children[0].children if hasattr(node.children[0], 'children') else node.children[0]
            self._add_highlighted_code(text_frame, str(content), lang)

        elif ntype == 'Table':
            if slide:
                self._add_table(slide, text_frame, node)

    def _fill_run(self, paragraph, node, is_title=False):
        if not hasattr(node, 'children'):
            return
        for child in node.children:
            ctype = child.__class__.__name__
            if isinstance(child, str):
                self._create_run(paragraph, child, is_title)
            elif ctype == 'RawText':
                self._create_run(paragraph, child.children, is_title)
            elif ctype == 'Strong':
                start_run = len(paragraph.runs)
                self._fill_run(paragraph, child, is_title)
                for i in range(start_run, len(paragraph.runs)):
                    paragraph.runs[i].font.bold = True
            elif ctype == 'Emphasis':
                start_run = len(paragraph.runs)
                self._fill_run(paragraph, child, is_title)
                for i in range(start_run, len(paragraph.runs)):
                    paragraph.runs[i].font.italic = True
            elif ctype == 'CodeSpan':
                self._create_run(paragraph, child.children, is_title, is_code=True)
            elif hasattr(child, 'children'):
                self._fill_run(paragraph, child, is_title)

    def _create_run(self, paragraph, text, is_title, bold=False, italic=False, is_code=False):
        if not text or not str(text).strip():
            return
        run = paragraph.add_run()
        run.text = str(text)
        run.font.name = self.code_font if is_code else self.font_name
        size = self.code_size if is_code else (self.title_size if is_title else self.body_size)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic

    # ---------------------- СОЗДАНИЕ СЛАЙДОВ -----------------
    def _create_title_slide(self, doc):
        slide = self.prs.slides.add_slide(self.title_layout)
        title_node, other_nodes = None, []
        for node in doc.children:
            if node.__class__.__name__ == 'Heading' and not title_node:
                title_node = node
            else:
                other_nodes.append(node)

        if slide.shapes.title and title_node:
            shape = self._position_shape(slide.shapes.title, top_inch=2.0, height_inch=1.5)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
            self._fill_run(p, title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=4.0, height_inch=2.5)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER)
            for node in other_nodes:
                self._add_node_to_frame(tf, node, slide=slide, default_align=PP_ALIGN.CENTER)

    def _create_content_slide(self, doc, slide_num):
        slide = self.prs.slides.add_slide(self.content_layout)
        title_node = None
        for node in doc.children:
            if node.__class__.__name__ == 'Heading':
                title_node = node
                break

        if slide.shapes.title and title_node:
            shape = self._position_shape(slide.shapes.title, top_inch=0.4, height_inch=1.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
            self._fill_run(p, title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=1.6, height_inch=5.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.LEFT)
            for node in doc.children:
                if node is title_node:
                    continue
                self._add_node_to_frame(tf, node, slide=slide)

    # ---------------------- ТОЧКИ ВХОДА ----------------------
    def create_from_text(self, md_text, output_path):
        self.warnings = []
        self.temp_files = []
        md_text = self._process_math_blocks(md_text)
        blocks = re.split(r'\n\s*---\s*\n', md_text.strip())
        if not blocks or not blocks[0].strip():
            raise PptxSyntaxError("MD текст пуст.")

        for idx, block in enumerate(blocks):
            if not block.strip():
                continue
            doc = self.md_parser.parse(block)
            if idx == 0:
                self._create_title_slide(doc)
            else:
                self._create_content_slide(doc, idx + 1)

        self.prs.save(output_path)
        for f in self.temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except:
                    pass
        return {"slides_created": len(self.prs.slides), "warnings": self.warnings}

    def create_from_file(self, md_path, output_path):
        with open(md_path, 'r', encoding='utf-8') as f:
            return self.create_from_text(f.read(), output_path)