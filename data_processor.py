"""
数据加载与预处理模块 - 通用文本分析
===================================
支持任意中文/英文文本数据的加载、清洗、标准化。
无任何行业特定逻辑，完全通用。

支持格式：xlsx, xls, csv, json, txt, pdf（自动识别）
"""

import os
import re
import json
import warnings
from pathlib import Path
from typing import Optional, Union, List, Dict, Any

import pandas as pd
import jieba
jieba.setLogLevel(jieba.logging.INFO)

warnings.filterwarnings('ignore')


# ══════════════════════════════════════════════════════════════
# 数据模型
# ══════════════════════════════════════════════════════════════

class Document:
    """单个文档"""

    def __init__(self, id: str, text: str,
                 attributes: Dict[str, Any] = None,
                 name: str = None):
        self.id = str(id)
        self.text = str(text) if text else ''
        self.name = name or f"文档_{id}"
        self.attributes = attributes or {}  # 任意属性：评分/日期/类型等
        self._segments = None
        self._word_list = None

    def __repr__(self):
        return f"<Document {self.id}: {self.name[:20]!r} ({len(self.text)}字)>"

    @property
    def segments(self) -> List[str]:
        """按汉语句子分隔符分割的段落列表（懒加载）"""
        if self._segments is None:
            self._segments = self._split_segments()
        return self._segments

    def _split_segments(self, min_len: int = 5) -> List[str]:
        """按中文句末标点分割句子"""
        parts = re.split(r'([。！？；\n])', self.text)
        segments = []
        buf = ''
        for part in parts:
            buf += part
            if part in '。！？；':
                stripped = buf.strip()
                if len(stripped) >= min_len:
                    segments.append(stripped)
                buf = ''
        if buf.strip():
            segments.append(buf.strip())
        return segments

    @property
    def word_list(self) -> List[str]:
        """分词列表（懒加载）"""
        if self._word_list is None:
            self._word_list = list(jieba.cut(self.text))
        return self._word_list

    def search(self, keyword: str, case_sensitive: bool = False) -> List[Dict]:
        """全文检索，返回匹配片段及上下文"""
        text = self.text if case_sensitive else self.text.lower()
        kw = keyword if case_sensitive else keyword.lower()
        results = []
        for seg in self.segments:
            s_text = seg if case_sensitive else seg.lower()
            idx = s_text.find(kw)
            if idx >= 0:
                results.append({
                    'document_id': self.id,
                    'document_name': self.name,
                    'segment': seg,
                    'position': idx,
                    'keyword': keyword,
                })
        return results

    def to_dict(self) -> dict:
        return {
            'id': self.id,
            'text': self.text,
            'name': self.name,
            'attributes': self.attributes,
        }

    @classmethod
    def from_dict(cls, d: dict) -> 'Document':
        return cls(
            id=d['id'],
            text=d['text'],
            name=d.get('name'),
            attributes=d.get('attributes', {}),
        )


class DocumentCollection:
    """
    文档集合 - 核心数据容器。
    支持按属性筛选、分组、统计。
    """

    def __init__(self, documents: List[Document] = None):
        self.documents: List[Document] = documents or []
        self._df = None

    def __len__(self):
        return len(self.documents)

    def __iter__(self):
        return iter(self.documents)

    def __getitem__(self, idx):
        return self.documents[idx]

    def add(self, doc: Document):
        self.documents.append(doc)
        self._df = None  # invalidate cache

    def add_collection(self, other: 'DocumentCollection'):
        """合并另一个集合"""
        self.documents.extend(other.documents)
        self._df = None

    @property
    def df(self) -> pd.DataFrame:
        """转换为 DataFrame（懒缓存）"""
        if self._df is None:
            rows = []
            for doc in self.documents:
                row = {'_id': doc.id, '_text': doc.text, '_name': doc.name}
                row.update(doc.attributes)
                rows.append(row)
            self._df = pd.DataFrame(rows)
        return self._df

    def filter(self, **attrs) -> 'DocumentCollection':
        """按属性筛选，如 dc.filter(评分=5)"""
        docs = []
        for doc in self.documents:
            match = all(doc.attributes.get(k) == v for k, v in attrs.items())
            if match:
                docs.append(doc)
        return DocumentCollection(docs)

    def group_by(self, attr: str) -> Dict[Any, 'DocumentCollection']:
        """按属性分组"""
        groups: Dict[Any, List[Document]] = {}
        for doc in self.documents:
            key = doc.attributes.get(attr, None)
            groups.setdefault(key, []).append(doc)
        return {k: DocumentCollection(v) for k, v in groups.items()}

    def stats(self) -> dict:
        """集合统计"""
        df = self.df
        text_cols = [c for c in df.columns if df[c].dtype == 'object']
        num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

        stats = {
            'total_documents': len(self),
            'total_characters': sum(len(d.text) for d in self.documents),
            'avg_document_length': sum(len(d.text) for d in self.documents) / max(len(self), 1),
            'attributes': list(df.columns),
            'numeric_attributes': num_cols,
            'text_attributes': text_cols,
        }

        # 数值属性统计
        for col in num_cols:
            stats[f'{col}_mean'] = df[col].mean()
            stats[f'{col}_median'] = df[col].median()
            stats[f'{col}_std'] = df[col].std()

        return stats

    def search(self, keyword: str, case_sensitive: bool = False,
               attrs: Dict[str, Any] = None) -> List[Dict]:
        """全文检索"""
        results = []
        for doc in self.documents:
            if attrs and not all(doc.attributes.get(k) == v for k, v in attrs.items()):
                continue
            for hit in doc.search(keyword, case_sensitive):
                results.append(hit)
        return results

    def save_json(self, path: str):
        data = [d.to_dict() for d in self.documents]
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @classmethod
    def load_json(cls, path: str) -> 'DocumentCollection':
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        docs = [Document.from_dict(d) for d in data]
        return cls(docs)


# ══════════════════════════════════════════════════════════════
# 文件加载
# ══════════════════════════════════════════════════════════════

# 文本列名的多语言别名映射
TEXT_COLUMN_ALIASES = [
    '评论内容', 'text', '内容', 'content', '文章', 'article',
    '正文', 'review', 'feedback', 'comment', 'description',
    '原始文本', 'full_text', 'review_text', 'comments',
]

NAME_COLUMN_ALIASES = [
    '文档名称', 'name', '标题', 'title', '文件名', 'doc_name',
    'document_name', '文档', 'subject', 'topic',
]

DATE_COLUMN_ALIASES = [
    '评论日期', 'date', '日期', 'created_at', 'publish_date',
    '时间', 'datetime', 'review_date', 'comment_date',
]

RATING_COLUMN_ALIASES = [
    '评分', 'rating', 'score', 'stars', 'rank',
    'rating_value',
]

CATEGORY_COLUMN_ALIASES = [
    '类型', 'category', 'type', 'tag', '标签', 'label',
    '分类', 'group', 'purpose',
]


def _find_column(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    """在 DataFrame 列名中查找匹配的别名"""
    cols = {str(c).strip(): c for c in df.columns}
    for alias in aliases:
        for col, orig in cols.items():
            if alias.lower() == col.lower() or alias in col:
                return str(orig)
    return None


def load_document(
    file_path: str,
    text_col: str = None,
    name_col: str = None,
    date_col: str = None,
    rating_col: str = None,
    category_col: str = None,
    id_col: str = None,
    encoding: str = 'utf-8',
) -> DocumentCollection:
    """
    加载单个文件为 DocumentCollection。

    Parameters
    ----------
    file_path : str
        文件路径（支持 xlsx/xls/csv/json/txt）
    text_col : str, optional
        文本列名（不指定则自动推断）
    name_col, date_col, rating_col, category_col : str, optional
        同上
    id_col : str, optional
        ID列（不指定则用行号）
    encoding : str
        txt文件编码，默认utf-8

    Returns
    -------
    DocumentCollection
    """
    path = Path(file_path)
    suffix = path.suffix.lower()

    # ── 加载原始 DataFrame ──────────────────────────────
    if suffix == '.txt':
        with open(path, 'r', encoding=encoding) as f:
            lines = [l.strip() for l in f if l.strip()]
        df = pd.DataFrame({'_text': lines})
        if id_col is None:
            df['_id'] = range(1, len(df) + 1)
    elif suffix == '.json':
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            df = pd.DataFrame(data)
        elif isinstance(data, dict):
            df = pd.DataFrame([data])
        else:
            raise ValueError(f"无法解析JSON格式: {type(data)}")
    elif suffix in ('.xlsx', '.xls'):
        try:
            df = pd.read_excel(path, engine='openpyxl' if suffix == '.xlsx' else 'xlrd')
        except Exception:
            df = pd.read_csv(path, encoding='utf-8-sig')
    elif suffix == '.csv':
        try:
            df = pd.read_csv(path, encoding='utf-8')
        except Exception:
            df = pd.read_csv(path, encoding='gbk')
    elif suffix == '.pdf':
        pages = extract_text_from_pdf(str(path))
        if not pages:
            raise ValueError("PDF 中未提取到文字（可能是扫描件或已加密）")
        df = pd.DataFrame({'_text': pages})
        df['_id'] = range(1, len(df) + 1)
    else:
        raise ValueError(f"不支持的文件格式: {suffix}")

    # ── 自动推断列名 ────────────────────────────────────
    if text_col is None:
        text_col = _find_column(df, TEXT_COLUMN_ALIASES)
    if name_col is None:
        name_col = _find_column(df, NAME_COLUMN_ALIASES)
    if date_col is None:
        date_col = _find_column(df, DATE_COLUMN_ALIASES)
    if rating_col is None:
        rating_col = _find_column(df, RATING_COLUMN_ALIASES)
    if category_col is None:
        category_col = _find_column(df, CATEGORY_COLUMN_ALIASES)
    if id_col is None:
        id_col = _find_column(df, ['id', 'ID', '编号', '序号'])

    # 默认文本列
    if text_col is None:
        # 取第一个非ID列
        for col in df.columns:
            if col not in ('_id',) and df[col].dtype == 'object':
                text_col = col
                break

    if text_col is None:
        raise ValueError(f"无法找到文本列，可用列: {list(df.columns)}")

    # ── 构造 DocumentCollection ─────────────────────────
    docs = []
    text_series = df[text_col].fillna('').astype(str)

    for i, (_, row) in enumerate(df.iterrows()):
        doc_id = str(row.get(id_col, i + 1)) if id_col else str(i + 1)
        text = str(row[text_col]) if text_col in df.columns else ''

        # 收集所有属性
        attrs = {}
        for col in df.columns:
            if col == text_col:
                continue
            val = row[col]
            if pd.isna(val):
                continue
            if isinstance(val, (int, float)):
                attrs[str(col)] = float(val)
            else:
                attrs[str(col)] = str(val)

        name = str(row[name_col]) if name_col and name_col in df.columns else f"文档_{doc_id}"
        docs.append(Document(id=doc_id, text=text, name=name, attributes=attrs))

    return DocumentCollection(docs)


def load_documents(
    *file_paths: str,
    text_col: str = None,
    name_col: str = None,
    **kwargs
) -> DocumentCollection:
    """
    批量加载多个文件。

    load_documents('file1.xlsx', 'file2.csv', text_col='评论内容')
    """
    all_docs = DocumentCollection()
    for fp in file_paths:
        dc = load_document(fp, text_col=text_col, name_col=name_col, **kwargs)
        all_docs.add_collection(dc)
    return all_docs


# ══════════════════════════════════════════════════════════════
# 文本预处理
# ══════════════════════════════════════════════════════════════

# 停用词
DEFAULT_STOPWORDS = set([
    '的', '了', '和', '是', '在', '就', '都', '而', '及', '与',
    '着', '或', '一个', '没有', '我们', '你们', '他们', '这个',
    '那个', '什么', '怎么', '如何', '为什么', '多少', '几个',
    '这', '那', '它', '她', '他', '自己', '自己', '自己',
    '也', '还', '很', '太', '真', '比较', '非常', '特别',
    '一些', '一下', '一点', '有些', '有的', '没', '不', '别',
    '要', '会', '能', '可以', '应该', '可能', '知道', '觉得',
    '啊', '呢', '吧', '呀', '嘛', '哦', '哈', '嗯', '噢',
    '但是', '可是', '如果', '因为', '所以', '虽然', '只是',
    '已经', '正在', '还有', '这些', '那些', '这样', '那样',
])


def preprocess_text(text: str,
                    remove_stopwords: bool = True,
                    remove_punctuation: bool = False) -> str:
    """
    单条文本预处理。

    Parameters
    ----------
    text : str
    remove_stopwords : bool  去除停用词
    remove_punctuation : bool  去除标点

    Returns
    -------
    str
    """
    import re
    if not isinstance(text, str):
        return ''

    # 去除URL
    text = re.sub(r'https?://\S+', '', text)
    # 去除邮箱
    text = re.sub(r'\S+@\S+', '', text)
    # 去除多余空格
    text = re.sub(r'\s+', ' ', text).strip()

    if remove_punctuation:
        import string
        punct = set(string.punctuation + '。，、；：？！""''（）【】《》…—～·')
        text = ''.join(ch for ch in text if ch not in punct)

    return text


def tokenize(text: str, stopwords: set = DEFAULT_STOPWORDS) -> List[str]:
    """分词（去除停用词）"""
    words = jieba.cut(text)
    return [w for w in words if w.strip() and w not in stopwords and len(w) > 1]


def basic_stats(collection: DocumentCollection) -> dict:
    """对 DocumentCollection 做基础统计"""
    df = collection.df
    text_col = '_text'
    num_cols = [c for c in df.columns
                if c not in ('_id', '_text', '_name')
                and pd.api.types.is_numeric_dtype(df[c])]

    stats = {
        'total_documents': len(collection),
        'total_characters': int(df[text_col].str.len().sum()),
        'avg_length': df[text_col].str.len().mean(),
        'min_length': df[text_col].str.len().min(),
        'max_length': df[text_col].str.len().max(),
    }

    for col in num_cols:
        stats[f'{col}_mean'] = float(df[col].mean())
        stats[f'{col}_median'] = float(df[col].median())
        stats[f'{col}_std'] = float(df[col].std())

    return stats


# ══════════════════════════════════════════════════════════════
# TF-IDF 文本分析
# ══════════════════════════════════════════════════════════════

def compute_tfidf(documents: List['Document'],
                  stopwords: set = None,
                  max_features: int = 500,
                  ngram_range: tuple = (1, 1)) -> 'pd.DataFrame':
    """
    计算 TF-IDF 矩阵。

    Parameters
    ----------
    documents : List[Document]
    stopwords : set  自定义停用词集合（与 DEFAULT_STOPWORDS 合并）
    max_features : int  最大词表规模
    ngram_range : tuple  (1,1)=unigram, (1,2)=+bigram, (1,3)=+trigram

    Returns
    -------
    pd.DataFrame: index=term, columns=doc_id, values=TF-IDF scores
    """
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
    except ImportError:
        raise ImportError("请安装 scikit-learn：pip install scikit-learn")

    sw = (stopwords or set()) | DEFAULT_STOPWORDS
    texts = [doc.text for doc in documents]
    ids = [doc.id for doc in documents]

    vectorizer = TfidfVectorizer(
        tokenizer=lambda t: [w for w in jieba.cut(t)
                              if w.strip() and w not in sw and len(w) > 1],
        max_features=max_features,
        ngram_range=ngram_range,
        min_df=1,
    )
    mat = vectorizer.fit_transform(texts)
    terms = vectorizer.get_feature_names_out()

    df = pd.DataFrame(mat.toarray(), index=ids, columns=terms).T
    df.index.name = 'term'
    return df


def compute_word_doc_matrix(documents: List['Document'],
                             stopwords: set = None,
                             top_n: int = 100) -> 'pd.DataFrame':
    """
    词-文档频次矩阵（原始词频）。

    Parameters
    ----------
    documents : List[Document]
    stopwords : set  自定义停用词
    top_n : int  仅返回频次最高的 top_n 个词

    Returns
    -------
    pd.DataFrame: index=term, columns=doc_id, values=词频计数
    """
    from collections import Counter
    sw = (stopwords or set()) | DEFAULT_STOPWORDS
    ids = [doc.id for doc in documents]
    all_words: Counter = Counter()
    per_doc: Dict[str, Counter] = {}

    for doc in documents:
        words = [w for w in jieba.cut(doc.text)
                 if w.strip() and w not in sw and len(w) > 1]
        c = Counter(words)
        per_doc[doc.id] = c
        all_words.update(c)

    top = [t for t, _ in all_words.most_common(top_n)]
    data = {did: {t: per_doc[did].get(t, 0) for t in top} for did in ids}
    return pd.DataFrame(data, index=top).T


# ══════════════════════════════════════════════════════════════
# PDF 文本提取
# ══════════════════════════════════════════════════════════════

def extract_text_from_pdf(file_path: str) -> List[str]:
    """
    从 PDF 文件提取文本，每页返回一条字符串。

    优先使用 pdfplumber（对中文支持更好），失败时回退到 PyPDF2。

    Parameters
    ----------
    file_path : str  PDF 文件路径

    Returns
    -------
    List[str]: 每页文本的列表
    """
    pages: List[str] = []

    # 尝试 pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        if pages:
            return pages
    except Exception:
        pass

    # 回退到 PyPDF2
    try:
        import PyPDF2
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    pages.append(text)
        return pages
    except ImportError:
        raise ImportError(
            "PDF 提取需要安装 pdfplumber 或 PyPDF2：\n"
            "  pip install pdfplumber\n或\n  pip install PyPDF2")
