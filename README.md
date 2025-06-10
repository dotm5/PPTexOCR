# PPTexOCR
PPTexOCR is a powerful batch processing tool designed to extract text and LaTeX formulas from PowerPoint (.pptx) slides. Combining OCR and advanced formula recognition, it enables efficient conversion of slide content into editable text, streamlining academic and professional workflows
# 简介
  这是一个基于 PySide6 GUI 的 PPTX 文件批量文字和公式识别工具，集成了 OCR（基于 pytesseract）和 LaTeX 公式识别（基于 pix2tex），支持：
- 批量添加 PPTX 文件
- 自动提取幻灯片中的文本框文本
- 对幻灯片中的图片进行文字与 LaTeX 公式识别
- 任务状态实时显示与进度条
- 结果导出为 TXT 文本文件
- 自定义圆角无边框半透明窗口，Windows 毛玻璃效果
- 支持拖拽添加文件，简洁美观的交互界面
# 环境及依赖
- Python 3.8 及以上版本
- PySide6 (GUI 框架)
- python-pptx (PPTX 文件解析)
- Pillow (图片处理)
- pytesseract (OCR)
- torch (深度学习框架)
- pix2tex (数学公式识别)
# 安装依赖
  建议使用虚拟环境安装：
```BASH
python -m venv venv
source venv/bin/activate       # Linux/macOS
venv\Scripts\activate.bat      # Windows
pip install -r requirements.txt
```
  Tesseract OCR 安装
pytesseract 是 Tesseract OCR 的 Python 封装，需要先安装 Tesseract OCR 工具：
- Windows: 下载并安装 https://github.com/tesseract-ocr/tesseract
- macOS: 使用 Homebrew brew install tesseract
- Linux: 通过 apt 或 yum 安装，如 sudo apt install tesseract-ocr
  安装后确保 tesseract 命令可用，并且将语言包（eng 和 chi_sim）安装完整。

# 使用说明
运行程序：
```BASH
python main.py
```
# 功能操作
- 添加PPT文件：点击按钮或拖拽 .pptx 文件至程序窗口
- 移除选中文件：从文件列表中选择后点击移除
- 开始识别：批量处理文件中的文本和图片OCR
- 导出选中文本：将选中文件识别结果保存为文本文件
# 界面特性
- 窗口无边框圆角设计，支持拖动
- Windows 平台启用毛玻璃半透明效果
- 进度条和日志显示识别状态
- 识别结果中包含纯文字和 LaTeX 公式文本
# 注意事项
  识别过程耗时，尤其包含图片公式识别，推荐使用 GPU 加速
  PPTX 文件必须是标准格式支持，非 PPT 文件不支持
  识别准确率依赖 Tesseract 和 pix2tex 模型
# 代码结构
- main.py : 主程序代码 (PySide6 GUI + OCR逻辑)
- requirements.txt : 依赖列表
- README.md : 使用说明及依赖介绍
# 贡献和反馈
欢迎提出 issues 或 PR，感谢使用本项目！
# 许可证
MIT License
