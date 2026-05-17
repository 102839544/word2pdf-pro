# Word2PDF Pro 🚀

**批量 Word 转 PDF 专业工具** - 界面精美、功能强大的商业软件

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)
![License](https://img.shields.io/badge/license-Commercial-red.svg)

---

## ✨ 功能特性

- 📄 **批量转换** - 支持同时选择多个 Word 文件批量转换为 PDF
- 🎨 **精美界面** - 基于 CustomTkinter 的现代化深色主题 UI
- 🏷️ **批量重命名** - 支持添加前缀、后缀、序号等多种重命名规则
- 🧹 **智能清理** - 自动去除 Word 文档中的批注、修订标记
- 🔐 **授权验证** - 机器码绑定授权系统，防止未授权使用
- 📊 **进度显示** - 实时显示转换进度，用户体验友好

---

## 📦 安装要求

### 系统要求
- Windows 10/11
- Python 3.8+
- Microsoft Word（用于批注清理功能）

### Python 依赖
```bash
pip install -r requirements.txt
```

依赖包：
- `customtkinter>=5.2.0` - 现代化 GUI
- `python-docx>=1.1.0` - Word 文档处理
- `docx2pdf>=0.1.8` - Word 转 PDF
- `pywin32>=306` - Windows COM 接口
- `Pillow>=10.0.0` - 图像处理

---

## 🚀 快速开始

### 1. 克隆项目
```bash
git clone https://github.com/102839544/word2pdf-pro.git
cd word2pdf-pro
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 运行程序
```bash
python src/app.py
```

### 4. 激活软件
首次运行需要输入授权码。授权码基于机器码生成，请联系开发商获取。

---

## 📖 使用指南

### 基本操作流程

1. **添加文件** - 点击"添加文件"按钮，选择要转换的 Word 文件（支持 .doc 和 .docx）
2. **设置选项** - 选择是否去除批注、设置重命名规则
3. **选择输出目录** - 点击"开始转换"后选择 PDF 文件的保存位置
4. **等待完成** - 进度条显示转换进度，完成后会弹出提示

### 重命名规则说明

- **不重命名** - 保持原文件名，仅修改扩展名为 .pdf
- **添加前缀** - 在原文件名前添加指定前缀
- **添加后缀** - 在原文件名后添加指定后缀
- **添加序号** - 按数字序号重命名（如：001.pdf, 002.pdf）

### 去除批注功能

勾选"去除批注和修订"后，软件会：
- 接受所有修订标记
- 删除所有批注
- 生成干净的 PDF 文件

---

## 🔐 授权系统

### 机器码生成
软件启动时会自动生成基于硬件信息的机器码（MD5 哈希的前16位）

### 授权码验证
授权码格式：32位大写字母数字组合
授权码与机器码绑定，无法在其他电脑上使用

### 获取授权
请联系：admin@example.com

---

## 📦 项目结构

```
word2pdf-pro/
├── src/
│   └── app.py              # 主程序（GUI + 核心功能）
├── tests/
│   └── test_convert.py     # 单元测试
├── docs/
│   └── user_manual.md      # 用户手册
├── requirements.txt         # Python 依赖
├── README.md               # 项目说明
├── LICENSE                 # 商业授权协议
└── .gitignore              # Git 忽略文件
```

---

## 🛠️ 技术栈

- **Python 3.8+** - 核心编程语言
- **CustomTkinter** - 现代化 GUI 框架
- **python-docx** - Word 文档读写
- **docx2pdf** - Word 转 PDF（基于 LibreOffice）
- **pywin32** - Windows COM 接口（处理批注）
- **Pillow** - 图像处理（未来功能预留）

---

## 📝 更新日志

### v1.0.0 (2026-05-17)
- ✅ 初始版本发布
- ✅ 批量 Word 转 PDF
- ✅ 批量重命名功能
- ✅ 去除批注和修订
- ✅ 授权验证系统
- ✅ 精美 GUI 界面

---

## ⚠️ 免责声明

1. 本软件为**商业软件**，未经授权不得分发、破解或修改
2. 使用本软件转换的文档，版权归原文档作者所有
3. 因使用本软件导致的任何数据丢失或损坏，开发商不承担责任
4. 请遵守当地法律法规，不要用于非法用途

---

## 📧 联系我们

- **技术支持**：admin@example.com
- **Bug 反馈**：https://github.com/102839544/word2pdf-pro/issues
- **GitHub**：https://github.com/102839544/word2pdf-pro

---

## ⭐ 支持项目

如果这个项目对你有帮助，欢迎：
- ⭐ 给项目点 Star
- 🐛 提交 Bug 报告
- 💡 提出功能建议
- 📢 分享给更多人

---

**© 2026 Word2PDF Pro. All rights reserved.**