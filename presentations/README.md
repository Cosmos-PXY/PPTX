# 生成可编辑 PPTX 的脚本与说明

本目录包含一个 Python 脚本（make_pptx_from_image.py），用于把你对话中上传的单页图片（命名为 input.png）自动转换为一页可编辑的 PPTX：每行 OCR 文本会被创建为独立文本框，底图为原图。

说明（快速流程）：
1. 在本目录上传你的图片并命名为 input.png（或在本地把图片保存为 input.png 并将其放在脚本同级目录）。
2. 在运行脚本前，请在运行环境中安装 Tesseract OCR（并安装中文语言包 chi_sim），以及 Python 包：python-pptx Pillow pytesseract opencv-python。
3. 运行脚本：
   python3 make_pptx_from_image.py
4. 脚本运行后会在同目录生成：研究现状_单页复刻.pptx

---

如果你希望我直接把最终的 PPTX 二进制文件也放到仓库（而不是脚本），请把 input.png 上传到仓库或允许我访问该图像数据；当前我无法直接从会话中读取二进制图片并把生成的 PPTX 直接推送为二进制文件到仓库，因此提供了脚本作为可重复、可审计的替代方案。
