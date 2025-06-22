## Section 在 MS Word 中的含义

Section（节） 是 MS Word 文档中的一个重要概念，它代表具有相同页面布局设置的文档部分。具体来说：

### 1. 页面布局设置

每个 Section 可以有不同的：* 页面方向：纵向（Portrait）或横向（Landscape）

* 页面尺寸：宽度和高度
* 页边距：上、下、左、右、装订线边距
* 页眉页脚距离

### 2. 节分隔符类型

Section 有不同的开始类型，对应 MS Word 中的节分隔符：* CONTINUOUS：连续节分隔符

* NEW_PAGE：下一页节分隔符
* NEW_COLUMN：分栏节分隔符
* EVEN_PAGE：偶数页节分隔符
* ODD_PAGE：奇数页节分隔符

### 3. 页眉页脚

每个 Section 可以定义自己的页眉页脚：* 默认页眉页脚

* 首页页眉页脚
* 奇偶页页眉页脚

### 4. 实际应用场景

在 MS Word 中，Section 常用于：* 混合页面方向：文档中部分页面纵向，部分横向

* 不同的页边距：不同部分使用不同的页边距设置
* 不同的页眉页脚：不同部分显示不同的页眉页脚内容
* 分栏布局：某些部分使用多栏布局

### 5. 默认情况

大多数 Word 文档默认只有一个 Section，除非用户手动插入节分隔符来创建新的节。

所以，document.py 中的 sections 属性返回的就是文档中所有的节，每个节对应 MS Word 中通过节分隔符分隔的文档部分。

# inline_shape

嵌入式形状是一种图形对象，例如图片，它包含在一段文本中，表现得像一个字符符号，与段落中的其他文本一样按顺序排列。

# Part 在 MS Word 中的含义

Part（部分） 是 Office Open XML 格式（包括 .docx 文件）中的一个核心概念，它基于 OPC（开放打包约定）标准。

## 1. OPC 架构基础

* OPC 是 Microsoft 开发的一种文件格式标准
* .docx 文件实际上是一个 ZIP 压缩包，包含多个 XML 文件
* 每个 XML 文件在 OPC 中称为一个 "Part"（部分）
  
## 2. DocumentPart 的具体含义

`DocumentPart` 是 Word 文档中的主文档部分，对应：**文件路径**：`/word/document.xml`

* 内容类型：`application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`
* 作用：包含文档的主要内容和结构

## 3. DocumentPart 管理的其他 Parts

`DocumentPart` 作为主文档部分，管理着文档中的其他部分：

### 核心部分：

* `/word/document.xml` - 主文档内容
* `/word/styles.xml` - 样式定义
* `/word/settings.xml` - 文档设置
* `/word/numbering.xml` - 编号定义
* `/word/comments.xml` - 批注内容

### 页眉页脚部分：

* `/word/header1.xml` - 页眉
* `/word/footer1.xml` - 页脚

### 文档属性：

* `/docProps/core.xml` - 核心属性（标题、作者、创建时间等）
* `/docProps/app.xml` - 应用程序属性

## 4. 在 MS Word 中的体现

在 MS Word 中，这些 Parts 对应：

### 用户可见的功能：

* 文档内容 - 用户编辑的文本、段落、表格等
* 样式 - 在"样式"面板中看到的样式定义
* 页面设置 - 在"页面设置"对话框中配置的选项
* 批注 - 在文档中显示的批注和修订
* 页眉页脚 - 在页面顶部和底部显示的内容

### 文档属性：

* 文件信息 - 在"文件 > 信息"中看到的文档属性
* 统计信息 - 字数、页数等统计信息

## 5. 技术架构
```txt
.docx 文件 (ZIP 包)
├── [Content_Types].xml          # 内容类型定义
├── _rels/.rels                  # 包级关系
├── word/
│   ├── document.xml             # 主文档部分 (DocumentPart)
│   ├── styles.xml               # 样式部分
│   ├── settings.xml             # 设置部分
│   ├── comments.xml             # 批注部分
│   ├── numbering.xml            # 编号部分
│   ├── header1.xml              # 页眉部分
│   └── footer1.xml              # 页脚部分
└── docProps/
    ├── core.xml                 # 核心属性部分
    └── app.xml                  # 应用程序属性部分
```

## 命名空间
```python
"""Constant values for OPC XML namespaces."""

DML_WORDPROCESSING_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
OFC_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OPC_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
OPC_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
WML_MAIN = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```