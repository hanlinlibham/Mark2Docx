import docx
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches, RGBColor, Pt
from docx.styles.style import _ParagraphStyle, BaseStyle
from docx.enum.style import WD_STYLE_TYPE
from docx.styles.styles import Styles

from .simple_style import SimpleStyle
from ..utils.style_enum import MDX_STYLE


class StyleManager:

    def __init__(self, doc: Document, yaml_conf: dict):
        self.document = doc
        self.style_conf = yaml_conf or {}
        # 默认样式配置
        self.default_style = {
            "h1": {
                "font": {
                    "size": 20,
                    "extra": ["bold"]
                }
            },
            "h2": {
                "font": {
                    "size": 16,
                    "extra": ["bold"]
                }
            },
            "h3": {
                "font": {
                    "size": 14,
                    "extra": ["bold"]
                }
            },
            "h4": {
                "font": {
                    "size": 12,
                    "extra": ["bold"]
                }
            }
        }
        # 合并默认样式和用户配置
        for key in self.default_style:
            if key not in self.style_conf:
                self.style_conf[key] = self.default_style[key]

        self.styles: Styles = doc.styles

    def init_styles(self):
        # 设置heading 1~4
        for i in range(1, 5):
            style_name = f"h{i}"
            if style_name in self.style_conf:
                s = SimpleStyle(f"Heading{i}", f"Heading {i}", self.style_conf[style_name])
                self.set_style(s)
        # TODO 还有什么样式设置呢？
        #  图片描述Caption、表格样式(?)

        # s = SimpleStyle("Normal", "Normal", self.style_conf["normal"])
        s = SimpleStyle(MDX_STYLE.PLAIN_TEXT, "Normal", self.style_conf["normal"])
        self.set_style(s)

    def set_style(self, _style: SimpleStyle):
        """设置样式"""
        style_id = self.styles.get_style_id(_style.style_name, _style.style_type)
        style_obj = None
        
        try:
            if style_id:
                style_obj = self.styles[_style.style_name]
            else:
                style_obj = self.styles.add_style(_style.style_name, _style.style_type)
        except ValueError:
            # 如果样式已存在，直接获取
            style_obj = self.styles[_style.style_name]

        # 设置字体
        font = style_obj.font
        font.name = _style.font_default
        # 将字体大小转换为 Word 的单位（1 pt = 2 * 12700 twips）
        font_size_twips = int(_style.font_size * 2 * 12700)
        font.size = Pt(_style.font_size)  # 这个设置可能不会生效
        # 直接设置 XML 元素的字体大小
        rPr = style_obj._element.get_or_add_rPr()
        sz = rPr.get_or_add_sz()
        sz.val = font_size_twips

        try:
            font.color.rgb = RGBColor(
                int(_style.font_color[0:2], 16),
                int(_style.font_color[2:4], 16),
                int(_style.font_color[4:6], 16)
            )
        except ValueError:
            print("[STYLE ERROR]: Invalid color format for", _style.style_name)

        # 设置字体附加属性
        font.bold = _style.font_bold
        font.italic = _style.font_italic
        font.underline = _style.font_underline
        font.strike = _style.font_strike

        # 设置段落格式
        if hasattr(style_obj, 'paragraph_format'):
            para_format = style_obj.paragraph_format
            para_format.first_line_indent = Pt(_style.first_line_indent)
            para_format.line_spacing = _style.line_spacing
            para_format.space_before = Pt(_style.space_before)
            para_format.space_after = Pt(_style.space_after)

        # 确保样式被应用
        if hasattr(style_obj, 'base_style'):
            style_obj.base_style = self.styles[_style.base_style_name]

        return style_obj

