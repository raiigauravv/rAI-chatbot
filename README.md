# 🤖 rAI bot - Universal File Converter & Multi-Document Analyzer

A **professional-grade ChatGPT-like file conversion platform** with **multi-document analysis capabilities**, built with Streamlit and powered by Databricks AI.

![Python](https://img.shields.io/badge/python-v3.8+-blue.svg)
![Streamlit](https://img.shields.io/badge/streamlit-v1.28+-red.svg)
![Status](https://img.shields.io/badge/status-active-success.svg)

## ✨ Key Features

### 🔄 **Professional File Conversion**
- **PDF → Word (DOCX)** - Advanced layout preservation with table detection
- **Multi-format support**: PDF, CSV, Excel, JSON, YAML, Markdown, HTML, TXT
- **High-quality conversions** using multiple extraction methods
- **Bulk conversion** to multiple formats simultaneously

### 📊 **Multi-Document Analysis** 
- **Analyze multiple documents together** with intelligent content combination
- **Cross-document comparison** and pattern detection
- **Smart keyword detection** for multi-document requests
- **Quick action buttons** for instant multi-document mode

### 🎯 **ChatGPT-like Interface**
- Clean, modern UI with ChatGPT-style conversations
- **Document switching** between multiple uploaded files  
- **Real-time processing** with visual feedback
- **Conversation history** with context preservation

## 🚀 Quick Start

### Prerequisites
- Python 3.8+
- Virtual environment (recommended)

### Installation

1. **Clone the repository**
```bash
git clone https://github.com/yourusername/rAI-bot-universal-converter.git
cd rAI-bot-universal-converter
```

2. **Create virtual environment**
```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run the application**
```bash
streamlit run app.py
```

5. **Open in browser**
- Local URL: http://localhost:8501
- The app will automatically open in your default browser

## 📋 Supported File Types

| Category | Formats | Notes |
|----------|---------|--------|
| **Documents** | PDF, DOCX, TXT, MD | Advanced PDF processing with layout preservation |
| **Data** | CSV, XLSX, JSON, YAML | Smart data structure detection |
| **Web** | HTML, CSS, JavaScript | Clean conversion with styling preservation |
| **Code** | Python (.py), JavaScript (.js) | Syntax-aware processing |

## 🎯 Usage Examples

### Single Document Conversion
```
"Convert this PDF to Word"
"Make this CSV into Excel format"
"Transform to JSON"
"Export as HTML"
```

### Multi-Document Analysis
```
"Compare all documents"
"Read both files together"  
"What are the differences between these files?"
"Analyze all my documents"
"Tell me about the other document"
```

### Advanced Features
```
"Convert to all formats" - Bulk conversion
"Compare all documents and find similarities"
"What patterns do you see across all files?"
```

## 🔧 Technology Stack

- **Frontend**: Streamlit (ChatGPT-like UI)
- **Backend**: Python 3.8+
- **AI Integration**: Databricks API
- **PDF Processing**: pdf2docx, PyMuPDF, pdfplumber
- **Data Processing**: pandas, openpyxl
- **Web Processing**: BeautifulSoup, html2text
- **Document Processing**: python-docx, markdown

## 📁 Project Structure

```
rAI-bot-universal-converter/
├── app.py                 # Main Streamlit application
├── chatbot_api.py        # Databricks API integration
├── test_agent.py         # Testing utilities
├── agent.py              # Core agent logic
├── requirements.txt      # Python dependencies
├── .gitignore           # Git ignore rules
└── README.md            # This file
```

## 🌟 Key Features in Detail

### Multi-Document Intelligence
- **Smart Detection**: Automatically detects when you want to analyze multiple files
- **Content Combination**: Merges documents with clear file separators
- **Context Optimization**: Truncates content intelligently to stay within API limits
- **Visual Feedback**: Shows exactly which files are being analyzed

### Conversion Quality
- **PDF → Word**: Professional-grade conversion preserving:
  - Tables and formatting
  - Layout structure  
  - Headers and sections
  - Images and graphics
- **Multiple extraction methods** with automatic fallback
- **Format-specific optimization** for each conversion type

### User Experience
- **ChatGPT-style interface** with conversation history
- **Document switching** with visual file management
- **Quick action buttons** for common operations
- **Real-time progress indicators**
- **Debug panel** for troubleshooting multi-document detection

# NoteGuardian 🛡️

A GitHub PR bot for Data Science and ML repos. Automatically reviews Jupyter notebooks and data files in pull requests, flags issues, and posts a single, always-up-to-date summary comment.

![Demo GIF](demo.gif) <!-- Replace with your actual demo GIF path -->

## Features
- 📝 **Notebook Output Checker:** Flags notebooks with cell outputs or execution counts.
- 📊 **Data File Detector:** Lists added/modified data files (csv, parquet, json, xlsx, feather, pkl).
- 📈 **Metrics Reporter:** If `metrics.json` is present, includes model metrics in the PR comment.
- 🔄 **Single Comment:** Updates its own comment to avoid PR noise.
- ⚡ **Fast & Lightweight:** Skips large notebooks for speed.

## Quickstart
1. **Copy these files to your repo:**
   - `.github/workflows/pr-bot.yml`
   - `scripts/pr_bot.py`
2. **Enable GitHub Actions** and set workflow permissions to "Read and write" in repo settings.
3. **Open a PR** with a notebook or data file change. The bot will comment automatically!

## Example PR Comment
```
### NoteGuardian 🛡️
_Analyzed PR #1 @ abc1234_

#### Notebooks changed
| File | Status |
|------|--------|
| `notebooks/demo_dirty.ipynb` | ⚠️ outputs present |

> Tip: Clear outputs via `jupyter nbconvert --ClearOutputPreprocessor.enabled=True --inplace your_notebook.ipynb`
> Or add a pre-commit hook: `nbstripout`

#### Data files changed
- `data.csv`

#### Model metrics
| Metric | Value |
|--------|-------|
| accuracy | 0.8723 |
| f1 | 0.8432 |
| precision | 0.8511 |
| recall | 0.8350 |
```

## One-click Setup
- Fork this repo or copy the workflow and script to your own.
- No extra secrets needed—uses GitHub’s built-in token.

## Why NoteGuardian?
- Saves reviewers time.
- Prevents accidental data leaks or dirty notebooks.
- Makes your repo look professional and recruiter-friendly.

## Contributing
PRs welcome! Ideas for more checks, file types, or integrations? Open an issue or PR.

## License
MIT

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with ❤️ using Streamlit
- Powered by Databricks AI
- Inspired by ChatGPT's user experience
- Professional PDF conversion capabilities

## 📧 Contact

- **Author**: Gaurav Rai
- **Project Link**: https://github.com/yourusername/rAI-bot-universal-converter

---

⭐ **Star this repo if you found it helpful!**

# Trigger bot update
# Trigger warning block test
