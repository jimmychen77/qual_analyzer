"""
QualCoder Pro - 通用质性分析工具
==================================
通用文本质性分析（QDA）桌面应用，参考 NVivo / ATLAS.ti / MAXQDA / QualCoder 设计，
完全通用、不包含任何行业特定内容。

核心功能：
  - 多文档管理（xlsx/xls/csv/json/txt，自动编码检测）
  - 层级编码系统（代码树，自由编码，自动编码）
  - 全文检索（关键词/正则，快速定位）
  - 交叉分析矩阵（编码 × 属性）
  - 多格式可视化图表
  - 备忘录系统（文档/代码/段落/项目级）
  - 多格式报告导出（Word/Excel/Markdown）
  - 项目保存/加载（.qda格式）

支持中文/英文文本，适用于社会科学、教育学、心理学、市场研究、政策分析等领域的质性研究。
"""

__version__ = '2.1.0'

# ── 路径修复 ─────────────────────────────────────────────────
# 确保无论从哪个目录运行 Python，hotel_analyzer 包都能被找到。
# Python 会自动把脚本所在目录加入 sys.path，
# 但运行 python hotel_analyzer/gui_app.py 时需要把 analysis/ 也加入。
import sys as _sys
from pathlib import Path as _Path

# 分析目录 = hotel_analyzer/ 的父目录
_ANALYSIS_DIR = _Path(__file__).resolve().parent.parent

# 加入 sys.path（去重）
if str(_ANALYSIS_DIR) not in _sys.path:
    _sys.path.insert(0, str(_ANALYSIS_DIR))

del _sys, _Path, _ANALYSIS_DIR

# ── 数据加载 ──────────────────────────────────────────────
from hotel_analyzer.data_processor import (
    Document,
    DocumentCollection,
    load_document,
    load_documents,
    basic_stats,
    compute_tfidf,
    compute_word_doc_matrix,
    extract_text_from_pdf,
)

# ── 编码系统 ──────────────────────────────────────────────
from hotel_analyzer.coding_browser import (
    Code,
    CodeSystem,
    ParagraphTagger,
    CrossTabAnalysis,
    CooccurrenceMatrix,
    AdvancedSearch,
    SegmentBrowser,
    CodeExporter,
)

# ── 备忘录 ──────────────────────────────────────────────
from hotel_analyzer.memo import (
    Memo,
    MemoManager,
)

# ── 分析函数 ──────────────────────────────────────────────
from hotel_analyzer.sentiment_analyzer import (
    SentimentIntensityAnalyzer,
    AspectSentimentAnalyzer,
    HiddenDissatisfactionDetector,
    KeywordAutoCoder,
    monthly_trend,
    infer_customer_persona,
    time_trend,
    segment_analysis,
)

# ── 可视化 ──────────────────────────────────────────────
from hotel_analyzer.visualizer import (
    Chart,
)

# ── 报告 ──────────────────────────────────────────────
from hotel_analyzer.reporter import (
    generate_word_report,
    generate_excel_report,
    generate_markdown_report,
)

# ── QDA 应用（顶层 API）─────────────────────────────────────────────
from hotel_analyzer.qda_app import (
    QDAApplication,
)

__all__ = [
    '__version__',
    'Document', 'DocumentCollection', 'load_document', 'load_documents',
    'infer_text_attribute', 'basic_stats',
    'Code', 'CodeSystem', 'ParagraphTagger',
    'CrossTabAnalysis', 'CooccurrenceMatrix', 'AdvancedSearch',
    'SegmentBrowser', 'CodeExporter',
    'Memo', 'MemoManager',
    'SentimentIntensityAnalyzer',
    'AspectSentimentAnalyzer',
    'HiddenDissatisfactionDetector',
    'KeywordAutoCoder',
    'monthly_trend', 'infer_customer_persona',
    'time_trend', 'segment_analysis',
    'Chart',
    'generate_word_report', 'generate_excel_report', 'generate_markdown_report',
    'QDAApplication',
]
