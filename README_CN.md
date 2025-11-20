# 文件转换器

一个功能强大的Dify插件，提供全面的本地文件转换功能。支持PDF、图像、Word、Excel、PowerPoint和文本格式之间的转换，具有高质量的输出和灵活的选项。

## 版本信息

- **当前版本**: v0.0.1
- **发布日期**: 2025-11-02
- **兼容性**: Dify Plugin Framework
- **Python版本**: 3.12

### 版本历史
- **v0.0.1** (2025-11-02): 初始版本，提供全面的文件转换功能

## 快速开始

1. 从Dify市场下载file_converter插件
2. 在您的Dify环境中安装插件
3. 立即开始使用各种转换工具

## 核心功能

### PDF转换
- **PDF转图片**: 将PDF文档转换为各种图像格式（.jpg、.jpeg、.png、.bmp、.tiff）
- **PDF转Word**: 从PDF提取内容并转换为Word文档
- **PDF转文本**: 从PDF文档提取文本内容

### 图像转换
- **图片转PDF**: 将一个或多个图像转换为PDF格式

### Microsoft Office转换
- **Word转PDF**: 将Word文档转换为PDF格式
- **Word转文本**: 从Word文档提取文本内容
- **Excel转PDF**: 将Excel电子表格转换为PDF格式
- **PowerPoint转PDF**: 将PowerPoint演示文稿转换为PDF格式

### 文本转换
- **文本转PDF**: 将纯文本文件转换为PDF格式
- **文本转Word**: 将纯文本文件转换为Word文档

## 技术优势

- **本地处理**: 所有转换都在本地执行，无需外部依赖
- **高质量输出**: 尽可能保持原始质量和格式
- **多格式支持**: 全面支持常见文件格式
- **错误处理**: 强大的错误处理，提供信息丰富的消息
- **灵活选项**: 针对不同用例的各种配置选项
- **安全处理**: 文件安全处理，不保留数据

## 要求

- Python 3.12
- Dify平台访问权限
- 所需的Python包（通过requirements.txt安装）

## 安装与配置

1. 安装所需依赖：
   ```bash
   pip install -r requirements.txt
   ```

2. 按照标准插件安装流程在您的Dify环境中安装插件

## 使用方法

该插件为不同的文件转换任务提供各种工具：

### PDF转换

#### 1. PDF转图片 (pdf_2_image)
将PDF文档转换为图像格式。
- **参数**:
  - `input_file`: 要转换的PDF文档（必填）
  - `output_format`: 输出图像格式（jpg、jpeg、png、bmp或tiff，默认：png）

#### 2. PDF转Word (pdf_2_word)
将PDF文档转换为Word格式。
- **参数**:
  - `input_file`: 要转换的PDF文档（必填）

#### 3. PDF转文本 (pdf_2_text)
从PDF文档提取文本内容。
- **参数**:
  - `input_file`: 要提取文本的PDF文档（必填）

### 图像转换

#### 4. 图片转PDF (image_2_pdf)
将一张或多张图片转换为单个PDF文档。
- **参数**:
  - `input_files`: 要转换的图片文件（必填，支持多文件）
- **功能特点**:
  - 支持多种图片格式（.jpg、.jpeg、.png、.bmp、.tiff）
  - 将多张图片转换为单个PDF文档
  - 保持图片上传顺序
  - 每张图片成为PDF中的一页

### Microsoft Office转换

#### 5. Word转PDF (word_2_pdf)
将Word文档转换为PDF格式。
- **参数**:
  - `input_file`: 要转换的Word文档（必填）

#### 6. Word转文本 (word_2_text)
从Word文档提取文本内容。
- **参数**:
  - `input_file`: 要提取文本的Word文档（必填）

#### 7. Excel转PDF (excel_2_pdf)
将Excel电子表格转换为PDF格式。
- **参数**:
  - `input_file`: 要转换的Excel文件（必填）

#### 8. PowerPoint转PDF (ppt_2_pdf)
将PowerPoint演示文稿转换为PDF格式。
- **参数**:
  - `input_file`: 要转换的PowerPoint文件（必填）

### 文本转换

#### 9. 文本转PDF (text_2_pdf)
将纯文本文件转换为PDF格式。
- **参数**:
  - `input_file`: 要转换的文本文件（必填）

#### 10. 文本转Word (text_2_word)
将纯文本文件转换为Word文档。
- **参数**:
  - `input_file`: 要转换的文本文件（必填）

## 注意事项

- 所有转换都在本地执行，无需将文件上传到外部服务
- 某些转换可能需要requirements.txt中包含的额外库
- 大文件可能需要更长的处理时间，具体取决于其复杂性和大小
- 输出文件的质量取决于输入文件的质量和格式

## 开发者信息

- **作者**: `https://github.com/sawyer-shi`
- **邮箱**: sawyer36@foxmail.com
- **许可证**: MIT License
- **源码地址**: `https://github.com/sawyer-shi/dify-plugins-file_converter`
- **支持**: 通过Dify平台和GitHub Issues

---

**准备好轻松转换您的文件了吗？**