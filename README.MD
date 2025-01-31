# MarkDocx

> 本项目是基于 [greyovo/markdocx](https://github.com/greyovo/markdocx) 的改进版本，感谢原作者的开源贡献。

一个强大的 Markdown 转 Word (.docx) 转换工具，支持丰富的 Markdown 语法和自定义样式。

## ✨ 特性

- 🎯 完整的 Markdown 语法支持
  - 标题（一到四级）与大纲导航
  - 列表（有序、无序、任务清单）
  - 表格（支持合并单元格）
  - 图片（本地和网络图片）
  - 文本格式（粗体、斜体、下划线、高亮等）
  - 超链接
  - 引用块
  - 代码块
  
- 🎨 灵活的样式定制
  - 通过 YAML 配置文件自定义样式
  - 支持中英文字体分别设置
  - 可调整段落格式（缩进、行距等）
  - 支持文本颜色和背景色

## 🚀 快速开始

### 安装

```bash
pip install markdocx
```

### 命令行使用

```bash
markdocx input.md -o output.docx -s style.yaml
```

参数说明：
- `input.md`: 输入的 Markdown 文件
- `-o output.docx`: 输出的 Word 文件路径（可选，默认与输入文件同名同目录）
- `-s style.yaml`: 自定义样式配置文件（可选，默认使用内置样式）
- `-a`: 转换完成后自动打开文件（可选）

### Python API 使用

```python
from markdocx import MarkDocx

# 初始化转换器
converter = MarkDocx(style_config="path/to/style.yaml")

# 转换文件
converter.convert("input.md", "output.docx", auto_open=True)

# 转换字符串
converter.convert_string("# Hello World", "output.docx")
```

## 🎨 样式配置

通过 YAML 文件自定义样式，示例：

```yaml
h1:  # 段落类型：h1-h4, normal
  font:
    default: Times New Roman  # 西文字体
    east-asia: 宋体          # 中文字体
    size: 12                 # 字体大小(pt)
    color: "000000"         # RGB颜色(16进制)
    extra:                  # 可选样式
      - bold               # 加粗
      - italic            # 斜体
      - underline         # 下划线
  first-line-indent: 0     # 首行缩进(倍)
  line-spacing: 1.2        # 行距(倍)
  space:
    before: 0              # 段前空格(pt)
    after: 0               # 段后空格(pt)
```

## 📝 示例

查看 `examples` 目录获取更多示例。

## 🔧 开发指南

1. 克隆仓库
```bash
git clone https://github.com/yourusername/markdocx.git
cd markdocx
```

2. 创建虚拟环境
```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows
```

3. 安装依赖
```bash
pip install -r requirements.txt
```

4. 运行测试
```bash
pytest tests/
```

## ⚠️ 已知限制

1. 行内样式限制：
   - 不支持多重行内样式（如同时加粗和斜体）
   - 列表最大支持两级嵌套

2. 暂不支持的功能：
   - LaTeX 公式
   - 脚注
   - 代码块语法高亮（计划支持）

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！

## 📄 开源协议

本项目采用 MIT 协议开源。

## 🙏 感谢

- [python-docx](https://python-docx.readthedocs.io)
- [python-markdown](https://python-markdown.github.io)
- [beautifulsoup4](https://beautifulsoup.readthedocs.io)
- [pyyaml](https://pyyaml.org)