# pdf-cropper

## 概述

本项目是一个基于 Python 的 GUI 工具，用于自动裁剪 PDF 文件中的图表，去除周围空白区域，同时允许用户设置需要忽略的边框宽度。该工具使用 `fitz` (PyMuPDF) 和 `PIL` (Pillow) 库来处理 PDF 文件和图像操作，并通过 `tkinter` 构建用户界面。

## 主要功能

- **自动裁剪**: 自动检测 PDF 页面中的内容区域并进行裁剪，去除多余的空白。
- **边框忽略**: 用户可自定义边框宽度，裁剪时忽略指定宽度的页面边缘。
- **批量处理**: 支持选择多个 PDF 文件进行批量处理。
- **用户友好的 GUI**: 基于 `tkinter` 的图形用户界面，操作简单直观。
- **DPI 感知**: 在 Windows 系统上启用 DPI 感知，确保在高分辨率屏幕上正常显示。
- **输出文件夹自定义**: 用户可以选择输出文件夹，或者使用默认的 `output` 文件夹。
- **处理后自动打开输出文件夹**: 在处理完成后，自动打开输出文件夹，方便用户查看处理结果。

## 环境要求

- **Python**: 3.11 或更高版本
- **依赖库**:
  - `tkinter`
  - `PyMuPDF` (`fitz`)
  - `Pillow` (`PIL`)
  - `numpy`

## 安装

1. 确保您已安装 Python 3.11 或更高版本。
2. 安装必要的 Python 库：
   ```bash
   pip install pymupdf Pillow numpy
   ```
3. 将项目代码下载或克隆到本地。

## 使用方法

1. 运行 `pdf_cropper.py` 脚本启动 GUI 程序。
2. **选择 PDF 文件**:
   - 点击 "选择文件" 按钮，选择要处理的 PDF 文件。支持多选。
   - 选择的文件路径将显示在文件列表框中。
3. **选择输出文件夹**:
   -  点击 “选择文件夹” 按钮，选择输出文件夹。如果未选择，将会在程序目录下生成一个名为`output`的文件夹。
4. **设置边框宽度**:
   - 在 “设置边框宽度” 输入框中输入需要忽略的边框宽度，默认为 5 像素。
5. **开始处理**:
   - 点击 “开始裁剪” 按钮开始处理。
   - 处理状态会在界面下方显示。
6. **查看结果**:
   - 处理完成后，会弹出完成提示，并自动打开输出文件夹。

## 贡献指南

欢迎提交 issues 和 pull requests! 如果您有任何改进建议或发现了 Bug，请及时告知。

1. Fork 本项目到您的 GitHub 账户。
2. 创建一个新的分支进行修改：`git checkout -b your-feature-branch`
3. 提交您的修改：`git commit -m "Add your feature or fix"`
4. 推送您的分支：`git push origin your-feature-branch`
5. 提交 pull request。

## 许可证

本项目使用 MIT 许可证，详情请查看 [LICENSE](LICENSE) 文件。

## 注意事项

- 如果你在使用过程中遇到任何问题，欢迎提交 Issue。

## 示例

![image](https://github.com/user-attachments/assets/f4788589-f836-44c5-a9df-f299c86e562c)


## 项目结构

```
.
├── pdf_cropper.py          # 主程序文件
├── README.md        # README 文件
├── LICENSE.md       # 许可证文件
```
