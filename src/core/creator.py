import os
import re
import uuid
import marko
from pathlib import Path
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Pt, Inches, Cm
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from copy import deepcopy


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

# Попытка импорта конвертера LaTeX -> Unicode (для простых инлайн формул)
try:
    from pylatexenc.latex2text import LatexNodes2Text
except ImportError:
    LatexNodes2Text = None

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

        # Новые настройки из .env
        self.tittle_margin_cm = float(os.getenv("PPT_TITTLE_MARGIN_CM", "0.0"))
        self.body_margin_cm = float(os.getenv("PPT_BODY_MARGIN_CM", "0.0"))
        self.title_bg_color = os.getenv("PPT_TITLE_BG_COLOR", "")
        self.title_font_color = os.getenv("PPT_TITLE_FONT_COLOR", "0,0,0")
        self.title_height_cm = float(os.getenv("PPT_TITLE_HEIGHT_CM", "1.5"))
        self.slide_numbering = os.getenv("PPT_SLIDE_NUMBERING", "false").lower() == "true"
        self.footer_text = os.getenv("PPT_FOOTER_TEXT", "")
        self.footer_height_cm = float(os.getenv("PPT_FOOTER_HEIGHT_CM", "1.0"))
        self.formula_numbering = os.getenv("PPT_FORMULA_NUMBERING", "false").lower() == "true"
        self.bullet_spacing = float(os.getenv("PPT_BULLET_SPACING", "12.0"))

        # Настройки для нижнего колонтитула и нумерации
        self.footer_font_size = int(os.getenv("PPT_FOOTER_FONT_SIZE", "12"))
        self.numbering_font_size = int(os.getenv("PPT_NUMBERING_FONT_SIZE", "14"))
        self.numbering_width_cm = float(os.getenv("PPT_NUMBERING_WIDTH_CM", "2.0"))

        self.formula_counter = 0

        # Загрузка новых параметров из env
        self.layout_title_idx = int(os.getenv("PPT_LAYOUT_TITLE_IDX", "0"))
        self.layout_content_idx = int(os.getenv("PPT_LAYOUT_CONTENT_IDX", "1"))
        self.content_bottom_buffer = float(os.getenv("PPT_CONTENT_BOTTOM_BUFFER_INCH", "1.2"))
        self.ts_title_top = float(os.getenv("PPT_TITLE_SLIDE_TITLE_TOP_INCH", "2.0"))
        self.ts_subtitle_top = float(os.getenv("PPT_TITLE_SLIDE_SUBTITLE_TOP_INCH", "4.0"))
        self.tf_padding = float(os.getenv("PPT_TEXT_FRAME_PADDING_CM", "0.13"))
        self.img_width_ratio = float(os.getenv("PPT_IMAGE_WIDTH_RATIO", "0.9"))
        self.table_row_h = float(os.getenv("PPT_TABLE_ROW_HEIGHT_CM", "1.0"))
        self.footer_border_color = os.getenv("PPT_FOOTER_BORDER_COLOR", "38,70,115")

        self._set_aspect_ratio(os.getenv("PPT_ASPECT_RATIO", "16:9"))
        self.title_layout = self.prs.slide_layouts[self.layout_title_idx]
        self.content_layout = self.prs.slide_layouts[self.layout_content_idx]

        self.warnings = []
        self.temp_files = []
        self.md_parser = marko.Markdown(extensions=['gfm'])

    def _parse_color(self, color_str, default_rgb=(0, 0, 0)):
        """Парсит цвет 'R,G,B' из строки."""
        if not color_str:
            return RGBColor(*default_rgb)
        try:
            parts = [int(x.strip()) for x in color_str.split(',')]
            return RGBColor(parts[0], parts[1], parts[2])
        except Exception:
            return RGBColor(*default_rgb)

    def _set_aspect_ratio(self, ratio_str):
        ratios = {"4:3": (9144000, 6858000), "16:9": (12192000, 6858000)}
        w, h = ratios.get(ratio_str, ratios["16:9"])
        self.prs.slide_width, self.prs.slide_height = w, h

    # ---------------------- КОНВЕРТАЦИЯ LaTeX В UNICODE -----------------
    def _latex_to_unicode(self, latex_str):
        """Преобразует простые LaTeX‑выражения в Unicode."""
        if not latex_str:
            return ""
        if LatexNodes2Text:
            try:
                converter = LatexNodes2Text()
                return converter.latex_to_text(latex_str)
            except Exception:
                pass

        replacements = {
            r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
            r'\epsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η', r'\theta': 'θ',
            r'\iota': 'ι', r'\kappa': 'κ', r'\lambda': 'λ', r'\mu': 'μ',
            r'\nu': 'ν', r'\xi': 'ξ', r'\pi': 'π', r'\rho': 'ρ',
            r'\sigma': 'σ', r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ',
            r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
            r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ',
            r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Upsilon': 'Υ',
            r'\Phi': 'Φ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
            r'\times': '×', r'\div': '÷', r'\pm': '±', r'\mp': '∓',
            r'\leq': '≤', r'\geq': '≥', r'\neq': '≠', r'\approx': '≈',
            r'\equiv': '≡', r'\propto': '∝', r'\infty': '∞',
            r'\int': '∫', r'\sum': '∑', r'\prod': '∏', r'\partial': '∂',
            r'\sqrt': '√', r'\nabla': '∇',
            r'^2': '²', r'^3': '³', r'^4': '⁴', r'^n': 'ⁿ',
            r'_0': '₀', r'_1': '₁', r'_2': '₂', r'_3': '₃', r'_i': 'ᵢ', r'_n': 'ₙ',
        }
        for pat, uni in replacements.items():
            latex_str = latex_str.replace(pat, uni)

        latex_str = re.sub(r'\\[a-zA-Z]+', '', latex_str)
        latex_str = latex_str.replace('{', '').replace('}', '')
        latex_str = latex_str.replace('$', '')
        return latex_str.strip()

    def _is_complex_formula(self, latex_str):
        """Определяет, является ли формула сложной (требует рендера в PNG)."""
        # Убрана триггерная проверка на одиночные _ и ^, чтобы простые E=mc^2 оставались текстом
        patterns = [
            r'\\frac', r'\\int', r'\\sum', r'\\prod', r'\\sqrt\{',
            r'\\lim', r'\\left', r'\\right', r'\\begin\{',
            r'_\{', r'\^\{'  # Сложные индексы и степени в фигурных скобках
        ]
        return any(re.search(p, latex_str) for p in patterns)

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
            fig.savefig(filename, transparent=True, bbox_inches='tight', pad_inches=0.01)
            plt.close(fig)
            self.temp_files.append(filename)
            return filename
        except Exception as e:
            self.warnings.append(f"Math error: {e}")
            plt.close(fig)
            return None

    def _process_math_blocks(self, text):
        """Обрабатывает блочные ($$) и инлайн ($) формулы.
        Блочные — всегда в PNG.
        Инлайн — если простая, в Unicode; если сложная — в PNG.
        """
        # Блочные формулы: заменяем на ![...](...)
        def block_replacer(m):
            img_path = self._render_formula_to_image(m.group(1))
            return f'\n![math]({img_path})\n' if img_path else m.group(0)

        text = re.sub(r'\$\$(.*?)\$\$', block_replacer, text, flags=re.DOTALL)

        # Инлайн формулы: проверяем сложность
        def inline_replacer(m):
            latex_expr = m.group(1).strip()
            if self._is_complex_formula(latex_expr):
                img_path = self._render_formula_to_image(latex_expr)
                return f'![math]({img_path})' if img_path else m.group(0)
            else:
                return self._latex_to_unicode(latex_expr)

        text = re.sub(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', inline_replacer, text)
        return text

    # ---------------------- ВСПОМОГАТЕЛЬНЫЕ -----------------
    def _apply_paragraph_style(self, paragraph, is_code=False, align=None, is_list_item=False):
        if align:
            paragraph.alignment = align
        paragraph.line_spacing = self.line_spacing

        if is_list_item:
            paragraph.space_after = Pt(self.bullet_spacing)
        else:
            self._remove_bullet_xml(paragraph)

    def _remove_bullet_xml(self, paragraph):
        """Удаляет маркер списка, но сохраняет отступы (level)."""
        pPr = paragraph._p.get_or_add_pPr()
        # Удаляем элемент маркера (a:buNone)
        for elem in pPr:
            if elem.tag.endswith('buNone'):
                pPr.remove(elem)
                break
        pPr.insert(0, OxmlElement('a:buNone'))
        # Не сбрасываем left_indent и first_line_indent, чтобы сохранить отступы списка

    def _setup_text_frame(self, shape, align=None, is_title=False):
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT if is_title else MSO_AUTO_SIZE.NONE
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if is_title else MSO_ANCHOR.TOP

        tf.margin_left = Cm(self.tittle_margin_cm) if is_title else Cm(self.body_margin_cm)
        tf.margin_right = Cm(self.tittle_margin_cm) if is_title else Cm(self.body_margin_cm)
        tf.margin_top = Cm(self.tf_padding)
        tf.margin_bottom = Cm(self.tf_padding)

        tf.clear()
        return tf

    def _position_shape(self, shape, top_inch, height_inch=None):
        # Ширина теперь: Ширина слайда минус (отступ * 2)
        slide_width = self.prs.slide_width
        new_width = slide_width - Cm(self.tittle_margin_cm * 2)
        shape.width = new_width
        if height_inch:
            shape.height = Inches(height_inch)

        shape.left = Cm(self.tittle_margin_cm)
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
                # Если это не технический микро-абзац под картинку (max_size=1), учитываем высоту
                if max_size > 1.0:
                    total_pt += max_size * 0.8
            else:
                chars_per_line = max(40, int(1400 / max_size))
                lines = (len(text) // chars_per_line) + 1
                total_pt += lines * max_size * self.line_spacing * 1.15

            if p.space_after:
                total_pt += p.space_after.pt

        return total_pt / 72.0

    # ---------------------- ВСТАВКА ИЗОБРАЖЕНИЙ И ТАБЛИЦ -----
    def _insert_image_shape(self, slide, text_frame, img_path, level=0, is_block=True, text_offset_inches=0):
        if not os.path.exists(img_path):
            return None

        current_text_h = self._get_current_tf_height(text_frame)
        offset_inch = 0.05 if is_block else -0.15
        top_inch = text_frame._parent.top.inches + current_text_h + offset_inch

        pic = slide.shapes.add_picture(img_path, Inches(0), Inches(top_inch))

        max_width = (self.prs.slide_width - Cm(self.tittle_margin_cm * 2)) * self.img_width_ratio
        if pic.width > max_width:
            ratio = max_width / pic.width
            pic.width = int(max_width)
            pic.height = int(pic.height * ratio)

        if is_block and level == 0:
            pic.left = int((self.prs.slide_width - pic.width) / 2)
        else:
            base_left = text_frame._parent.left
            indent_margin = Inches(0.4 * level + 0.1)
            pic.left = base_left + indent_margin + Inches(text_offset_inches)

        # Нумерация блочных формул
        is_math = "math_" in img_path
        if is_block and is_math and self.formula_numbering:
            self.formula_counter += 1
            num_w = Cm(2)
            num_l = self.prs.slide_width - Cm(self.tittle_margin_cm) - num_w
            num_shape = slide.shapes.add_textbox(num_l, Inches(top_inch), num_w, pic.height)
            tf = num_shape.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            p.text = f"({self.formula_counter})"
            p.font.size = Pt(self.body_size)
            p.font.name = self.font_name

        if is_block:
            p = text_frame.add_paragraph()
            self._remove_bullet_xml(p)
            p.text = " "
            p.font.size = Pt(1)
            p.space_before = Pt(pic.height.pt + 10)

        return pic

    def _add_table(self, slide, text_frame, node):
        rows_nodes = node.children
        rows = len(rows_nodes)
        cols = max((len(r.children) for r in rows_nodes), default=0)
        if rows == 0 or cols == 0:
            return

        current_text_h = self._get_current_tf_height(text_frame)
        top_inch = text_frame._parent.top.inches + current_text_h + 0.1

        width = int(self.prs.slide_width * self.img_width_ratio)
        left = int((self.prs.slide_width - width) / 2)
        row_height_cm = self.table_row_h
        tbl_shape = slide.shapes.add_table(rows, cols, left, Inches(top_inch), width, Cm(row_height_cm * rows))
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
    def _add_node_to_frame(self, text_frame, node, slide=None, level=0, default_align=None, is_list_item=False):
        ntype = node.__class__.__name__

        if ntype == 'Paragraph':
            p = self._get_or_add_paragraph(text_frame)
            p.level = min(level, 8)
            self._apply_paragraph_style(p, align=default_align, is_list_item=is_list_item)

            children = getattr(node, 'children', [])
            is_block = len(children) == 1

            for child in children:
                if child.__class__.__name__ == 'Image':
                    if slide:
                        text_before = "".join(r.text for r in p.runs)
                        text_offset = len(text_before) * 0.13

                        pic = self._insert_image_shape(
                            slide, text_frame, child.dest,
                            level=p.level, is_block=is_block, text_offset_inches=text_offset
                        )

                        if not is_block and pic:
                            run = p.add_run()
                            space_count = max(1, int(pic.width.inches * 14))
                            run.text = " " * space_count
                else:
                    self._fill_run(p, child)

        elif ntype == 'List':
            for item in node.children:
                if item.__class__.__name__ == 'ListItem':
                    for sub in item.children:
                        if sub.__class__.__name__ == 'List':
                            self._add_node_to_frame(text_frame, sub, slide=slide, level=level + 1)
                        else:
                            self._add_node_to_frame(text_frame, sub, slide=slide, level=level,
                                                    default_align=default_align, is_list_item=True)

        elif ntype in ['FencedCode', 'CodeBlock']:
            lang = getattr(node, 'lang', 'text')
            content = node.children[0].children if hasattr(node.children[0], 'children') else node.children[0]
            self._add_highlighted_code(text_frame, str(content), lang)

        elif ntype == 'Table':
            if slide:
                self._add_table(slide, text_frame, node)

    def _fill_run(self, paragraph, node, is_title=False, bold=False, italic=False):
        """Рекурсивно обходит AST и добавляет текст с накопленными стилями."""
        if node is None:
            return

        ntype = node.__class__.__name__

        # ИСПРАВЛЕНИЕ: marko использует имя StrongEmphasis для жирного шрифта
        current_bold = bold or (ntype in ['Strong', 'StrongEmphasis'])
        current_italic = italic or (ntype == 'Emphasis')
        is_code = (ntype == 'CodeSpan')

        # 1) Узел — обычная строка (например, при прямом вызове)
        if isinstance(node, str):
            self._create_run(paragraph, node, is_title, current_bold, current_italic, is_code)
            return

        # 2) Сырой текст (`RawText`)
        if ntype == 'RawText':
            self._create_run(paragraph, node.children, is_title, current_bold, current_italic, is_code)
            return

        # 3) Узел имеет атрибут `children` (список дочерних узлов)
        if hasattr(node, 'children'):
            children = node.children
            # Защита: иногда `children` может быть строкой
            if isinstance(children, str):
                self._create_run(paragraph, children, is_title, current_bold, current_italic, is_code)
                return
            # Рекурсивно обрабатываем всех детей
            for child in children:
                self._fill_run(paragraph, child, is_title, current_bold, current_italic)
            return

        # 4) Узел имеет атрибут `text` (например, `LineBreak`)
        if hasattr(node, 'text'):
            self._create_run(paragraph, node.text, is_title, current_bold, current_italic, is_code)
            return

        # 5) Запасной вариант — просто преобразуем узел в строку
        text = str(node)
        if text:
            self._create_run(paragraph, text, is_title, current_bold, current_italic, is_code)

    def _create_run(self, paragraph, text, is_title, bold=False, italic=False, is_code=False):
        """Создаёт Run и применяет накопленные свойства."""
        if text is None or not str(text):
            return

        run = paragraph.add_run()
        run.text = str(text)

        # Шрифт и размер
        run.font.name = self.code_font if is_code else self.font_name
        size = self.code_size if is_code else (self.title_size if is_title else self.body_size)
        run.font.size = Pt(size)

        if is_title:
            run.font.color.rgb = self._parse_color(self.title_font_color, (0, 0, 0))

        # Жирность и курсив
        run.font.bold = bold or is_title
        run.font.italic = italic

    # ---------------------- СОЗДАНИЕ СЛАЙДОВ -----------------
    def _add_footer_and_numbering(self, slide, slide_idx):
        """Отрисовывает нижний колонтитул и номер слайда по макету."""
        slide_w = self.prs.slide_width
        slide_h = self.prs.slide_height
        margin = Cm(self.tittle_margin_cm)
        footer_h = Cm(self.footer_height_cm)

        bottom_y = slide_h - footer_h

        border_color = self._parse_color(self.footer_border_color, (38, 70, 115))

        num_w = Cm(self.numbering_width_cm) if self.slide_numbering else 0

        if self.footer_text:
            # Ширина футера точно рассчитывается, чтобы оставить место под номер
            footer_w = slide_w - margin * 2 - num_w

            f_shape = slide.shapes.add_textbox(margin, bottom_y, footer_w, footer_h)
            f_shape.line.color.rgb = border_color
            f_shape.line.width = Pt(1)
            tf = f_shape.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = self.footer_text
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(self.footer_font_size)
            p.font.name = self.font_name

        if self.slide_numbering:
            # Левая координата стыкуется ровно с правым краем футера
            num_l = margin + footer_w if self.footer_text else slide_w - margin - num_w

            n_shape = slide.shapes.add_textbox(num_l, bottom_y, num_w, footer_h)
            n_shape.name = "SlideNumberBox"  # Добавляем имя, чтобы легко найти на втором проходе
            n_shape.line.color.rgb = border_color
            n_shape.line.width = Pt(1)
            tf = n_shape.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = str(slide_idx)  # Временно пишем только номер
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(self.numbering_font_size)
            p.font.bold = True
            p.font.name = self.font_name

    def _init_content_slide(self, title_node, slide_idx, is_continuation=False):
        """Инициализирует новый контентный слайд с правильными стилями и фонами."""
        slide = self.prs.slides.add_slide(self.content_layout)

        if slide.shapes.title and title_node:
            shape = self._position_shape(slide.shapes.title, top_inch=Cm(self.tittle_margin_cm).inches,
                                         height_inch=Cm(self.title_height_cm).inches)

            # Применение фона заголовка
            if self.title_bg_color:
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = self._parse_color(self.title_bg_color, (255, 255, 255))

            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)

            # Если это перенос, добавляем (продолжение)
            if is_continuation:
                copied_node = deepcopy(title_node)
                if hasattr(copied_node.children[0], 'children'):
                    copied_node.children[0].children += " (продолжение)"
                self._fill_run(p, copied_node, is_title=True)
            else:
                self._fill_run(p, title_node, is_title=True)

        self._add_footer_and_numbering(slide, slide_idx)

        # Вычисляем стартовую позицию для текста (отступ сверху от заголовка)
        top_offset_inch = Cm(self.tittle_margin_cm + self.title_height_cm).inches
        tf = None
        if len(slide.placeholders) > 1:
            # ВАЖНО: Считаем доступную высоту для блока текста
            footer_h_inch = Cm(self.footer_height_cm).inches if (self.footer_text or self.slide_numbering) else Cm(
                self.tittle_margin_cm).inches
            content_h_inch = self.prs.slide_height.inches - top_offset_inch - footer_h_inch - Cm(self.tittle_margin_cm).inches

            # Теперь передаем height_inch в _position_shape
            shape = self._position_shape(slide.placeholders[1], top_inch=top_offset_inch, height_inch=content_h_inch)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.LEFT, is_title=False)

        return slide, tf

    def _create_title_slide(self, doc, slide_num):
        slide = self.prs.slides.add_slide(self.title_layout)
        title_node, other_nodes = None, []
        for node in doc.children:
            if node.__class__.__name__ == 'Heading' and not title_node:
                title_node = node
            else:
                other_nodes.append(node)

        if slide.shapes.title and title_node:
            shape = self._position_shape(slide.shapes.title, top_inch=self.ts_title_top, height_inch=Cm(self.title_height_cm).inches)
            if self.title_bg_color:
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = self._parse_color(self.title_bg_color, (255, 255, 255))

            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER, is_title=True)
            p = tf.paragraphs[0]
            self._apply_paragraph_style(p, align=PP_ALIGN.CENTER)
            self._fill_run(p, title_node, is_title=True)

        if len(slide.placeholders) > 1:
            shape = self._position_shape(slide.placeholders[1], top_inch=self.ts_subtitle_top, height_inch=2.5)
            tf = self._setup_text_frame(shape, align=PP_ALIGN.CENTER)
            for node in other_nodes:
                self._add_node_to_frame(tf, node, slide=slide, default_align=PP_ALIGN.CENTER)

        self._add_footer_and_numbering(slide, slide_num)

    def _create_content_slide(self, doc, start_slide_idx):
        title_node = next((n for n in doc.children if n.__class__.__name__ == 'Heading'), None)
        nodes_to_process = [n for n in doc.children if n is not title_node]

        slide_idx = start_slide_idx
        slide, tf = self._init_content_slide(title_node, slide_idx)

        # Вычисляем доступную высоту (Auto-splitting logic)
        footer_h = Cm(self.footer_height_cm).inches if self.footer_text or self.slide_numbering else 0
        title_h = Cm(self.title_height_cm + self.tittle_margin_cm).inches
        # Оставляем ~1.2 дюйма буфера снизу, чтобы текст не наезжал на колонтитул
        max_h_inches = self.prs.slide_height.inches - title_h - footer_h - self.content_bottom_buffer

        for node in nodes_to_process:
            # Проверяем высоту перед вставкой (защита от переполнения)
            if tf and self._get_current_tf_height(tf) > max_h_inches and len(tf.paragraphs) > 1:
                slide_idx += 1
                slide, tf = self._init_content_slide(title_node, slide_idx, is_continuation=True)

            if tf:
                self._add_node_to_frame(tf, node, slide=slide)

        return slide_idx

    # ---------------------- ТОЧКИ ВХОДА ----------------------
    def create_from_text(self, md_text, output_path):
        self.warnings = []
        self.temp_files = []
        self.formula_counter = 0  # Сброс нумерации для нового файла

        md_text = self._process_math_blocks(md_text)
        blocks = re.split(r'\n\s*---\s*\n', md_text.strip())
        if not blocks or not blocks[0].strip():
            raise PptxSyntaxError("MD текст пуст.")

        slide_num = 1
        for idx, block in enumerate(blocks):
            if not block.strip():
                continue
            doc = self.md_parser.parse(block)
            if idx == 0:
                self._create_title_slide(doc, slide_num)
                slide_num += 1
            else:
                slide_num = self._create_content_slide(doc, slide_num) + 1

        # --- ВТОРОЙ ПРОХОД: Добавление общего количества слайдов ---
        total_slides = len(self.prs.slides)
        if self.slide_numbering:
            for slide in self.prs.slides:
                for shape in slide.shapes:
                    if shape.name == "SlideNumberBox":
                        tf = shape.text_frame
                        if tf.paragraphs:
                            p = tf.paragraphs[0]
                            current_num = p.text
                            p.text = f"{current_num}/{total_slides}"
                            # Переназначаем стили, так как перезапись p.text их сбрасывает
                            p.alignment = PP_ALIGN.CENTER
                            p.font.size = Pt(self.numbering_font_size)
                            p.font.bold = True
                            p.font.name = self.font_name

        self.prs.save(output_path)
        for f in self.temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except:
                    pass
        return {"slides_created": total_slides, "warnings": self.warnings}

    def create_from_file(self, md_path, output_path):
        with open(md_path, 'r', encoding='utf-8') as f:
            return self.create_from_text(f.read(), output_path)