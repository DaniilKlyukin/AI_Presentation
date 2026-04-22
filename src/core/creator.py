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

# Для рендера формул
try:
    import matplotlib.pyplot as plt
except ImportError:
    plt = None

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
        self.temp_files = []  # Хранилище временных файлов (формул/картинок)
        self.md_parser = marko.Markdown(extensions=['gfm'])

    def _set_aspect_ratio(self, ratio_str):
        ratios = {"4:3": (9144000, 6858000), "16:9": (12192000, 6858000)}
        w, h = ratios.get(ratio_str, ratios["16:9"])
        self.prs.slide_width, self.prs.slide_height = w, h

    # ----- ЛОГИКА ДЛЯ ФОРМУЛ -----
    def _render_formula_to_image(self, formula_text):
        """Рендерит LaTeX формулу в PNG с прозрачным фоном."""
        if not plt:
            self.warnings.append("Установите matplotlib для рендера формул (pip install matplotlib)")
            return None

        formula = formula_text.strip()
        filename = f"math_tmp_{uuid.uuid4().hex}.png"

        # Настраиваем matplotlib
        fig = plt.figure(figsize=(0.01, 0.01))
        # Рендерим текст как формулу ($...$)
        fig.text(0, 0, f"${formula}$", fontsize=24, color='black', ha='center', va='center')

        try:
            fig.savefig(filename, format='png', transparent=True, bbox_inches='tight', pad_inches=0.1)
            plt.close(fig)
            return filename
        except Exception as e:
            self.warnings.append(f"Ошибка рендера формулы: {e}")
            plt.close(fig)
            return None

    def _process_math_blocks(self, text):
        """Находит $$ формулы $$ и заменяет их на маркдаун картинки."""
        math_blocks = re.findall(r'\$\$(.*?)\$\$', text, flags=re.DOTALL)
        for math in math_blocks:
            img_path = self._render_formula_to_image(math)
            if img_path:
                self.temp_files.append(img_path)
                # Заменяем блок формулы на стандартный тег изображения Markdown
                text = text.replace(f"$${math}$$", f"![math]({img_path})")
        return text

    def _insert_image_shape(self, slide, text_frame, img_path):
        """Вставляет картинку на слайд, вычисляя отступ от текущего текста."""
        if not os.path.exists(img_path): return

        # Вычисляем, где мы находимся по вертикали
        text_lines = sum(1 for p in text_frame.paragraphs if p.text.strip())
        top_inch = 1.6 + (text_lines * 0.35)
        top_inch = min(top_inch, 5.5)

        # Добавляем изображение
        pic = slide.shapes.add_picture(img_path, Inches(0), Inches(top_inch))

        # Центрируем картинку по горизонтали
        pic.left = int((self.prs.slide_width - pic.width) / 2)

        # Чтобы следующий текст не налез на картинку, добавляем пустые строки
        empty_lines_needed = int((pic.height.inches / 0.35) + 1)
        for _ in range(empty_lines_needed):
            self._get_or_add_paragraph(text_frame)

    # -----------------------------

    def _apply_paragraph_style(self, paragraph, is_code=False, align=None):
        if align: paragraph.alignment = align
        paragraph.line_spacing = self.line_spacing

        if is_code:
            paragraph.level = 0
            pPr = paragraph._p.get_or_add_pPr()
            from pptx.oxml.xmlchemy import OxmlElement
            buNone = OxmlElement('a:buNone')
            pPr.insert(0, buNone)
            paragraph.left_indent = 0
            paragraph.first_line_indent = 0
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)

    def _setup_text_frame(self, shape, align=None, is_title=False):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if is_title else MSO_ANCHOR.TOP
        tf.clear()
        return tf

    def _position_shape(self, shape, top_inch, height_inch, width_percent=0.9):
        slide_width = self.prs.slide_width
        new_width = int(slide_width * width_percent)
        shape.width, shape.height = new_width, Inches(height_inch)
        shape.left, shape.top = int((slide_width - new_width) / 2), Inches(top_inch)
        return shape

    def _get_or_add_paragraph(self, text_frame):
        if len(text_frame.paragraphs) == 1 and not text_frame.paragraphs[0].text:
            return text_frame.paragraphs[0]
        return text_frame.add_paragraph()

    def _add_highlighted_code(self, text_frame, code_text, lang='text'):
        code_text = code_text.replace('\t', '    ').strip('\n\r')
        try:
            lexer = get_lexer_by_name(lang if lang else 'text')
        except:
            lexer = lexers.get_lexer_by_name('text')

        lines = code_text.split('\n')
        for i, line in enumerate(lines):
            if i == 0:
                p = self._get_or_add_paragraph(text_frame)
            else:
                p = text_frame.add_paragraph()
            self._apply_paragraph_style(p, is_code=True, align=PP_ALIGN.LEFT)
            if not line: continue

            tokens = lexer.get_tokens(line)
            for ttype, value in tokens:
                clean_value = value.replace('\r', '').replace('\n', '')
                if not clean_value: continue

                run = p.add_run()
                run.text = clean_value
                run.font.name = self.code_font
                run.font.size = Pt(self.code_size)

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

    def _add_node_to_frame(self, text_frame, node, slide=None, level=0, default_align=None):
        ntype = node.__class__.__name__

        if ntype == 'Paragraph':
            # Проверяем, есть ли внутри параграфа картинка (или наша отрендеренная формула)
            contains_image = any(c.__class__.__name__ == 'Image' for c in getattr(node, 'children', []))
            if contains_image:
                for child in node.children:
                    if child.__class__.__name__ == 'Image':
                        self._insert_image_shape(slide, text_frame, child.dest)
                return  # Прерываем, так как картинка обработана как отдельный блок

            p = self._get_or_add_paragraph(text_frame)
            p.level = min(level, 8)
            self._apply_paragraph_style(p, align=default_align)
            self._fill_run(p, node)

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
            if not slide: return
            rows, cols = len(node.children), max((len(r.children) for r in node.children), default=0)
            if rows == 0 or cols == 0: return

            text_lines = sum(1 for p in text_frame.paragraphs if p.text.strip())
            top_inch = min(1.6 + (text_lines * 0.35), 5.0)

            width = int(self.prs.slide_width * 0.9)
            left = int((self.prs.slide_width - width) / 2)
            table_shape = slide.shapes.add_table(rows, cols, left, Inches(top_inch), width, Inches(0.5 * rows))
            table = table_shape.table

            for r_idx, row_node in enumerate(node.children):
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
                            if r_idx == 0: run.font.bold = True

            # Добавляем пустые строки под таблицей
            empty_lines_needed = int(rows * 1.5)
            for _ in range(empty_lines_needed):
                self._get_or_add_paragraph(text_frame)

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
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
            if title_node: self._fill_run(p, title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=4.2, height_inch=2.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER)
            for node in other_nodes:
                self._add_node_to_frame(tf, node, slide=slide, default_align=PP_ALIGN.CENTER)

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
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
            if title_node: self._fill_run(p, title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=1.6, height_inch=5.0)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.LEFT)
            for node in content_nodes:
                self._add_node_to_frame(tf, node, slide=slide)

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

        if code:
            run.font.name = self.code_font
            run.font.size = Pt(self.code_size)
        else:
            run.font.name = self.font_name
            run.font.size = Pt(self.title_size if is_title else self.body_size)

        run.font.bold, run.font.italic = bold, italic

    def create_from_text(self, md_text, output_path):
        self.warnings = []
        self.temp_files = []  # Сбрасываем временные файлы

        # 1. Заменяем $$ формулы $$ на картинки
        md_text = self._process_math_blocks(md_text)

        blocks = re.split(r'\n\s*---\s*\n', md_text.strip())
        if not blocks or not blocks[0].strip(): raise PptxSyntaxError("MD текст пуст.")

        for idx, block in enumerate(blocks):
            if not block.strip(): continue
            doc = self.md_parser.parse(block)
            if idx == 0:
                self._create_title_slide(doc)
            else:
                self._create_content_slide(doc, idx + 1)

        self.prs.save(output_path)

        # Очистка мусора: удаляем сгенерированные PNG картинки формул
        for tmp_file in self.temp_files:
            if os.path.exists(tmp_file):
                try:
                    os.remove(tmp_file)
                except:
                    pass

        return {"slides_created": len(self.prs.slides), "warnings": self.warnings}

    def create_from_file(self, md_path, output_path):
        with open(md_path, 'r', encoding='utf-8') as f: return self.create_from_text(f.read(), output_path)