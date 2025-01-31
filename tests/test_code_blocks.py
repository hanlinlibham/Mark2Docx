import pytest
from pathlib import Path
from markdocx import MarkDocx
import tempfile
import os
import shutil

# 设置测试文件路径
TEST_FILE = Path(__file__).parent.parent / "example" / "code_block_test.md"

@pytest.fixture(scope="session", autouse=True)
def setup_test_environment():
    """设置测试环境"""
    # 创建示例目录
    example_dir = Path(__file__).parent.parent / "example"
    example_dir.mkdir(exist_ok=True)
    
    # 创建配置目录
    config_dir = Path(__file__).parent.parent / "markdocx" / "config"
    config_dir.mkdir(exist_ok=True)
    
    # 确保默认样式配置文件存在
    default_style = config_dir / "default_style.yaml"
    if not default_style.exists():
        default_style.write_text("""
# 默认样式配置
normal:
  font:
    default: Times New Roman
    east-asia: 宋体
    size: 12
    color: "000000"
  line-spacing: 1.3
  space:
    before: 7
    after: 7

heading_1:
  font:
    default: 黑体
    east-asia: 黑体
    size: 22
  line-spacing: 1.2
  space:
    before: 11
    after: 11
""")
    
    # 确保测试文件存在
    if not TEST_FILE.exists():
        original_file = Path(__file__).parent.parent.parent / "example" / "code_block_test.md"
        if original_file.exists():
            shutil.copy2(original_file, TEST_FILE)
        else:
            TEST_FILE.write_text("""
# Test Document

This is a test document for code blocks.

```cpp
int main() {
    cout << "hello world" << endl;
    return 0;
}
```

> Some quote text

1. First item
2. Second item
   - Sub item 1
   - Sub item 2
""")

def test_code_block_conversion():
    # 初始化转换器
    try:
        converter = MarkDocx()
    except Exception as e:
        pytest.fail(f"Failed to initialize MarkDocx: {e}")
    
    # 使用临时文件进行测试
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        output_path = Path(tmp.name)
        
        try:
            # 转换文档
            result = converter.convert(
                str(TEST_FILE),
                output_path
            )
            
            # 验证文件是否创建成功
            assert result.exists()
            assert result.stat().st_size > 0
            
        except Exception as e:
            pytest.fail(f"Conversion failed: {e}")
        finally:
            # 清理临时文件
            if output_path.exists():
                output_path.unlink()

def test_code_block_with_progress():
    # 记录进度回调的调用
    progress_calls = []
    
    def progress_callback(progress: float, message: str):
        progress_calls.append((progress, message))
    
    # 初始化转换器
    converter = MarkDocx()
    converter.set_progress_callback(progress_callback)
    
    # 使用临时文件进行测试
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        output_path = Path(tmp.name)
        
        try:
            # 获取测试文件的绝对路径
            test_file = Path(__file__).parent.parent / "example" / "code_block_test.md"
            
            # 转换文档
            converter.convert(
                str(TEST_FILE),
                output_path
            )
            
            # 验证进度回调
            assert len(progress_calls) > 0
            # 验证最后一个进度是100%
            assert progress_calls[-1][0] == 1.0
            assert "completed" in progress_calls[-1][1].lower()
            
        finally:
            # 清理临时文件
            if output_path.exists():
                output_path.unlink()

def test_code_block_with_custom_style():
    # 自定义样式配置
    style_config = {
        "normal": {
            "font": {
                "default": "Consolas",
                "east-asia": "微软雅黑",
                "size": 10,
                "color": "000000"
            },
            "line-spacing": 1.15
        }
    }
    
    # 初始化转换器
    converter = MarkDocx(style_config=style_config)
    
    # 使用临时文件进行测试
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        output_path = Path(tmp.name)
        
        try:
            # 获取测试文件的绝对路径
            test_file = Path(__file__).parent.parent / "example" / "code_block_test.md"
            
            # 转换文档
            result = converter.convert(
                str(TEST_FILE),
                output_path
            )
            
            # 验证文件是否创建成功
            assert result.exists()
            assert result.stat().st_size > 0
            
        finally:
            # 清理临时文件
            if output_path.exists():
                output_path.unlink()

if __name__ == "__main__":
    pytest.main([__file__]) 