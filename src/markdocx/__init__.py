"""
MarkDocx - Convert Markdown to Word documents with customizable styles
"""

from pathlib import Path
import yaml
from yaml import FullLoader

from .parser.md_parser import md2html
from .provider.docx_processor import DocxProcessor


class MarkDocx:
    def __init__(self, style_config=None):
        """
        初始化MarkDocx转换器
        
        Args:
            style_config: 样式配置，可以是YAML文件路径或字典
        """
        self.style_config = {}
        if style_config:
            if isinstance(style_config, str):
                with open(style_config, 'r', encoding='utf-8') as f:
                    self.style_config = yaml.load(f, FullLoader)
            elif isinstance(style_config, dict):
                self.style_config = style_config
            else:
                raise ValueError("style_config must be a file path or dictionary")
        else:
            # 使用默认样式
            default_style = Path(__file__).parent / "config" / "default_style.yaml"
            if default_style.exists():
                with open(default_style, 'r', encoding='utf-8') as f:
                    self.style_config = yaml.load(f, FullLoader)
        
        self.processor = DocxProcessor(style_conf=self.style_config)
    
    def convert(self, input_path, output_path=None, auto_open=False):
        """
        转换Markdown文件到Word文档
        
        Args:
            input_path: Markdown文件路径
            output_path: 输出Word文件路径，默认为输入文件同目录
            auto_open: 是否自动打开生成的文件
            
        Returns:
            输出文件路径
        """
        input_path = Path(input_path)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        if output_path is None:
            output_path = input_path.with_suffix('.docx')
        else:
            output_path = Path(output_path)
        
        # 转换过程
        html_path = input_path.with_suffix('.html')
        md2html(str(input_path), str(html_path))
        
        try:
            self.processor.html2docx(str(html_path), str(output_path))
            if auto_open and output_path.exists():
                import os
                os.startfile(str(output_path.absolute()))
            return output_path
        finally:
            # 清理临时HTML文件
            if html_path.exists():
                html_path.unlink()
    
    def convert_string(self, markdown_string, output_path):
        """
        转换Markdown字符串到Word文档
        
        Args:
            markdown_string: Markdown内容
            output_path: 输出Word文件路径
            
        Returns:
            输出文件路径
        """
        temp_md = Path('temp.md')
        try:
            temp_md.write_text(markdown_string, encoding='utf-8')
            return self.convert(temp_md, output_path)
        finally:
            if temp_md.exists():
                temp_md.unlink() 