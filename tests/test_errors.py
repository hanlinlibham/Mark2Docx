import pytest
from pathlib import Path
from markdocx import MarkDocx
from markdocx.utils.exceptions import FileError, ConfigError, ParserError, ProcessorError
from . import create_temp_file, cleanup_temp_file

def test_file_not_found():
    """测试文件不存在的情况"""
    converter = MarkDocx()
    with pytest.raises(FileError, match="Input file not found"):
        converter.convert("nonexistent.md")

def test_invalid_style_config_file():
    """测试无效的样式配置文件"""
    with pytest.raises(ConfigError, match="Invalid YAML format"):
        MarkDocx(style_config="nonexistent_style.yaml")

def test_invalid_style_config_content():
    """测试无效的样式配置内容"""
    invalid_yaml = create_temp_file("invalid:yaml:content", ".yaml")
    try:
        with pytest.raises(ConfigError):
            MarkDocx(style_config=str(invalid_yaml))
    finally:
        cleanup_temp_file(invalid_yaml)

def test_invalid_markdown():
    """测试无效的Markdown内容"""
    converter = MarkDocx()
    with pytest.raises(ParserError):
        converter.convert_string("```unclosed_code_block", "output.docx")

def test_invalid_output_path():
    """测试无效的输出路径"""
    converter = MarkDocx()
    with pytest.raises(FileError):
        converter.convert("input.md", "/nonexistent/directory/output.docx")

def test_empty_markdown():
    """测试空Markdown文件"""
    empty_file = create_temp_file("", ".md")
    try:
        converter = MarkDocx()
        with pytest.raises(ParserError, match="Empty markdown content"):
            converter.convert(str(empty_file))
    finally:
        cleanup_temp_file(empty_file)

def test_invalid_image_path():
    """测试无效的图片路径"""
    md_content = "![Invalid Image](nonexistent.png)"
    md_file = create_temp_file(md_content, ".md")
    try:
        converter = MarkDocx()
        with pytest.raises(ProcessorError, match="Image file not found"):
            converter.convert(str(md_file))
    finally:
        cleanup_temp_file(md_file)

def test_invalid_style_values():
    """测试无效的样式值"""
    invalid_style = {
        "normal": {
            "font": {
                "size": "invalid_size",  # 应该是数字
                "color": "invalid_color"  # 应该是有效的颜色代码
            }
        }
    }
    with pytest.raises(ConfigError, match="Invalid style value"):
        MarkDocx(style_config=invalid_style)

def test_recursive_markdown_include():
    """测试递归包含的Markdown文件"""
    content1 = "# Test\n{{include.md}}"
    content2 = "## Include\n{{test.md}}"
    
    file1 = create_temp_file(content1, ".md")
    file2 = create_temp_file(content2, ".md")
    
    try:
        converter = MarkDocx()
        with pytest.raises(ParserError, match="Recursive include detected"):
            converter.convert(str(file1))
    finally:
        cleanup_temp_file(file1)
        cleanup_temp_file(file2)

def test_large_file_handling():
    """测试大文件处理"""
    # 创建一个超过限制的大文件
    large_content = "# " + "x" * 1024 * 1024 * 10  # 10MB的内容
    large_file = create_temp_file(large_content, ".md")
    
    try:
        converter = MarkDocx()
        with pytest.raises(ProcessorError, match="File too large"):
            converter.convert(str(large_file))
    finally:
        cleanup_temp_file(large_file) 