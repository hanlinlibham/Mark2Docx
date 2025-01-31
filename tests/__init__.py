"""
Test package for markdocx.
Contains test utilities and fixtures.
"""

import pytest
from pathlib import Path
import tempfile
import shutil

# 测试资源目录
TEST_RESOURCES = Path(__file__).parent / "resources"

def get_test_file(filename: str) -> Path:
    """获取测试文件路径"""
    return TEST_RESOURCES / filename

def create_temp_file(content: str, suffix: str = ".md") -> Path:
    """创建临时测试文件"""
    temp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    temp.write(content.encode('utf-8'))
    temp.close()
    return Path(temp.name)

def cleanup_temp_file(path: Path):
    """清理临时文件"""
    if path.exists():
        path.unlink()

@pytest.fixture(scope="session", autouse=True)
def setup_test_resources():
    """设置测试资源目录"""
    TEST_RESOURCES.mkdir(exist_ok=True)
    
    # 创建测试用的markdown文件
    test_md = TEST_RESOURCES / "test.md"
    if not test_md.exists():
        test_md.write_text('''
# Test Document

## Text Formatting
This is **bold**, *italic*, and `code`.

## Lists
1. First item
2. Second item
   - Sub item 1
   - Sub item 2

## Code Block
```python
def hello():
    print("Hello, World!")
```

## Table
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |

## Image
![Test Image](test.png)

## Link
[Example Link](https://example.com)
''')
    
    # 创建测试用的图片文件
    test_img = TEST_RESOURCES / "test.png"
    if not test_img.exists():
        # 创建一个简单的测试图片
        shutil.copy(Path(__file__).parent.parent / "assets" / "example.png", test_img)
    
    yield
    
    # 清理临时文件
    # shutil.rmtree(TEST_RESOURCES)  # 如果需要清理，取消注释 