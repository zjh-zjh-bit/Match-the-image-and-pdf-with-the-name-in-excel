# 证书文件智能分类工具

一个智能的证书文件分类工具，能够根据Excel表格中的学生信息，自动识别和分类图片及PDF证书文件。

## 功能特点

-  **识别**：支持文件名匹配和OCR文字识别
-  **多格式支持**：处理JPG、PNG、PDF等多种格式
-  **Excel集成**：直接从Excel表格读取学生信息
-  **一键打包**：生成完全独立的可执行文件
-  **无需安装**：所有依赖内置，开箱即用

## 快速开始

### 方法一：使用预编译版本
1. 从 [here](https://github.com/zjh-zjh-bit/Match-the-image-and-pdf-with-the-name-in-excel/releases/download/v1.0.0/default.exe) 下载最新版本
2. 解压后运行 `证书分类工具.exe`
3. 按照界面提示操作

证书分类工具 - 发布说明

版本: v1.0.0
发布日期: 2025/11/01
作者: zjh

发布内容:
- 主程序: 证书分类工具.exe
- 依赖文件: Tesseract OCR + Poppler
- 使用说明文档

主要功能:
1. 从Excel读取学生信息
2. 智能识别图片和PDF中的姓名
3. 自动分类和复制匹配文件
4. 支持OCR文字识别

文件列表:
- 证书分类工具.exe (主程序)
- 使用说明.txt (详细使用指南)
- poppler/ (PDF处理工具)
- tessdata/ (OCR识别引擎)

系统要求:
- Windows 10/11 64位
- 无需安装Python
- 建议安装VC++运行时

已知问题:
- 无
