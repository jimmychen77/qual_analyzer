"""
QDA 应用核心类
===================================
参考 NVivo / ATLAS.ti / MAXQDA / QualCoder 的设计理念，
实现通用的质性数据分析应用。

核心功能：
- 多文档管理（DocumentCollection）
- 层级编码系统（CodeSystem）
- 备忘录（MemoManager）
- 交叉分析矩阵（CrossTabAnalysis）
- 共现分析（CooccurrenceMatrix）
- 全文检索（AdvancedSearch）
- 可视化（Chart系列）
- 报告导出（Word/Excel/Markdown）
"""

import os
import json
import zipfile
import base64
from pathlib import Path
from typing import Dict, List, Any, Optional, Union

import pandas as pd

from hotel_analyzer.data_processor import (
    Document, DocumentCollection,
    load_document, load_documents,
    basic_stats,
)
from hotel_analyzer.coding_browser import (
    Code, CodeSystem,
    ParagraphTagger, CrossTabAnalysis, CooccurrenceMatrix,
    AdvancedSearch, SegmentBrowser, CodeExporter,
)
from hotel_analyzer.memo import Memo, MemoManager
from hotel_analyzer.visualizer import Chart
from hotel_analyzer.sentiment_analyzer import (
    SentimentIntensityAnalyzer,
    AspectSentimentAnalyzer,
    HiddenDissatisfactionDetector,
    KeywordAutoCoder,
    monthly_trend,
    infer_customer_persona,
)


class QDAApplication:
    """
    QDA应用主类。

    设计参考 QualCoder / Taguette 的开源架构，
    完全通用，不含任何行业特定内容。

    示例用法::

        app = QDAApplication()

        # 加载文档
        app.load_documents('reviews.xlsx', text_col='评论内容')

        # 全文检索
        hits = app.search('服务质量')

        # 创建编码
        app.create_code('服务态度', color='#1976d2', description='前台/客服态度')
        app.assign_code(doc_id='1', segment='很好', code_name='服务态度')

        # 交叉分析
        matrix = app.cross_tab('评分分组')

        # 导出
        app.export_report('output.docx')
    """

    VERSION = '2.0.0'

    def __init__(self, project_name: str = '未命名项目'):
        self.project_name = project_name
        self.project_path: Optional[str] = None

        # ── 核心数据 ──────────────────────────────────
        self.documents: DocumentCollection = DocumentCollection()
        self.code_system: CodeSystem = CodeSystem()
        self.memo_manager: MemoManager = MemoManager()

        # ── 分析缓存 ──────────────────────────────────
        self._analysis_results: Dict[str, Any] = {}

        # ── 配色方案（QualCoder风格）───────────────
        self.CODE_COLORS = [
            '#1a73e8', '#d93025', '#1e8e3e', '#9334e6',
            '#f5511e', '#007b83', '#f9ab00', '#e91e63',
            '#1565c0', '#ff5722', '#00bcd4', '#795548',
        ]

        # ── 分析器 ──────────────────────────────────
        self.sentiment_analyzer = SentimentIntensityAnalyzer()
        self.aspect_analyzer = AspectSentimentAnalyzer()
        self.hidden_neg_analyzer = HiddenDissatisfactionDetector()

    # ════════════════════════════════════════════════════
    # 文档管理
    # ════════════════════════════════════════════════════

    def load_documents(self, *file_paths: str,
                      text_col: str = None,
                      name_col: str = None,
                      **kwargs) -> DocumentCollection:
        """
        加载文档集合。

        Parameters
        ----------
        file_paths : str
            多个文件路径
        text_col : str, optional
            文本列名（不指定则自动推断）
        name_col : str, optional
            文档名列名（不指定则自动推断）

        Returns
        -------
        DocumentCollection
        """
        self.documents = load_documents(*file_paths, text_col=text_col, name_col=name_col, **kwargs)
        self._analysis_results.clear()
        return self.documents

    def add_document(self, text: str, name: str = None,
                    **attributes) -> Document:
        """添加单个文档"""
        doc_id = str(len(self.documents) + 1)
        doc = Document(id=doc_id, text=text, name=name or f"文档_{doc_id}",
                      attributes=attributes)
        self.documents.add(doc)
        return doc

    def get_document(self, doc_id: str) -> Optional[Document]:
        """按ID获取文档"""
        for doc in self.documents:
            if doc.id == doc_id:
                return doc
        return None

    def document_stats(self) -> dict:
        """文档集合统计"""
        return basic_stats(self.documents)

    # ════════════════════════════════════════════════════
    # 编码系统
    # ════════════════════════════════════════════════════

    def create_code(self, name: str,
                   color: str = None,
                   description: str = '',
                   parent_name: str = None) -> Code:
        """
        创建新编码。

        Parameters
        ----------
        name : str  编码名称
        color : str  十六进制颜色，如 '#1a73e8'
        description : str  描述
        parent_name : str, optional  父编码名称（用于层级）

        Returns
        -------
        Code
        """
        if color is None:
            used_colors = {c.color for c in self.code_system.all_codes.values()}
            for cc in self.CODE_COLORS:
                if cc not in used_colors:
                    color = cc
                    break
            else:
                color = self.CODE_COLORS[len(self.code_system.all_codes) % len(self.CODE_COLORS)]

        code = self.code_system.add_code(
            name=name,
            color=color,
            description=description,
            parent_name=parent_name,
        )
        return code

    def assign_code(self, doc_id: str,
                    segment: str,
                    code_name: str,
                    memo: str = '') -> bool:
        """
        为文档片段分配编码。

        Parameters
        ----------
        doc_id : str  文档ID
        segment : str  匹配的文本片段
        code_name : str  编码名称
        memo : str, optional  批注

        Returns
        -------
        bool 是否成功
        """
        doc = self.get_document(doc_id)
        if doc is None:
            return False

        instance = {
            'doc_id': doc_id,
            'segment': segment,
            'start': doc.text.find(segment) if segment in doc.text else -1,
            'end': -1,
            'memo': memo,
        }

        # all_codes 以 code.id（UUID）为 key，不能用 code_name（字符串）作 existence check
        # 必须通过 name 查找；不存在则创建编码，再添加实例
        if not self.code_system._find_code_by_name(code_name):
            self.create_code(code_name)

        self.code_system.add_instance(code_name, instance)
        return True

    def remove_code(self, code_name: str) -> bool:
        """删除编码（同时删除所有实例）"""
        return self.code_system.remove_code(code_name)

    def get_code_instances(self, code_name: str) -> List[Dict]:
        """获取编码的所有片段"""
        code = self.code_system.all_codes.get(code_name)
        if code is None:
            return []
        return list(code.instances)

    def auto_code_from_keywords(self,
                                keyword_dict: Dict[str, List[str]],
                                doc_id: str = None,
                                match_whole_word: bool = False) -> Dict[str, int]:
        """
        基于关键词自动编码。

        Parameters
        ----------
        keyword_dict : dict  {编码名: [关键词列表]}
        doc_id : str, optional  限定文档（None表示全部）
        match_whole_word : bool  全词匹配

        Returns
        -------
        dict  {编码名: 匹配数量}
        """
        import re
        results = {}

        docs = [self.get_document(doc_id)] if doc_id else list(self.documents)

        for code_name, keywords in keyword_dict.items():
            count = 0
            # 自动创建编码（通过名称查找，避免all_codes的int key问题）
            if not self.code_system._find_code_by_name(code_name):
                self.create_code(code_name)

            for doc in docs:
                if doc is None:
                    continue
                for kw in keywords:
                    if match_whole_word:
                        pattern = rf'\b{re.escape(kw)}\b'
                        hits = re.findall(pattern, doc.text)
                    else:
                        hits = [m.start() for m in re.finditer(re.escape(kw), doc.text)]

                    for _ in hits:
                        inst = {
                            'doc_id': doc.id,
                            'segment': kw,
                            'start': -1,
                            'end': -1,
                            'memo': '自动编码',
                        }
                        self.code_system.add_instance(code_name, inst)
                        count += 1

            results[code_name] = count

        return results

    # ════════════════════════════════════════════════════
    # 备忘录
    # ════════════════════════════════════════════════════

    def add_document_memo(self, doc_id: str, text: str,
                          memo_type: str = 'general') -> Memo:
        """为文档添加备忘录"""
        return self.memo_manager.add_doc_memo(doc_id, text, memo_type)

    def add_code_memo(self, code_name: str, text: str,
                     memo_type: str = 'general') -> Memo:
        """为编码添加备忘录"""
        if not self.code_system._find_code_by_name(code_name):
            self.create_code(code_name)
        code = self.code_system._find_code_by_name(code_name)
        return self.memo_manager.add_code_memo(code.id, text, memo_type)

    def add_project_memo(self, text: str,
                         memo_type: str = 'general') -> Memo:
        """添加项目级备忘录"""
        return self.memo_manager.add_project_memo(text, memo_type)

    def get_memos(self, memo_type: str = None,
                  linked_id: str = None) -> List[Memo]:
        """获取备忘录"""
        return self.memo_manager.filter_memos(memo_type=memo_type, linked_id=linked_id)

    # ════════════════════════════════════════════════════
    # 全文检索
    # ════════════════════════════════════════════════════

    def search_fulltext(self, keyword: str,
                       case_sensitive: bool = False,
                       doc_id: str = None,
                       attrs: Dict[str, Any] = None) -> List[Dict]:
        """
        全文检索。

        Returns
        -------
        list of dict: {doc_id, doc_name, segment, position, keyword}
        """
        if doc_id:
            doc = self.get_document(doc_id)
            return doc.search(keyword, case_sensitive) if doc else []

        return self.documents.search(keyword, case_sensitive, attrs)

    def advanced_search(self, pattern: str,
                        use_regex: bool = True) -> List[Dict]:
        """
        高级检索（支持正则）。

        Returns
        -------
        list of dict: {doc_id, doc_name, match, position}
        """
        analyzer = AdvancedSearch()
        docs = list(self.documents)
        return analyzer.search_documents(docs, pattern, use_regex)

    # ════════════════════════════════════════════════════
    # 交叉分析
    # ════════════════════════════════════════════════════

    def build_cross_matrix(self,
                          attribute: str,
                          top_n: int = 10) -> pd.DataFrame:
        """
        构建编码 × 属性的交叉分析表。

        Parameters
        ----------
        attribute : str  属性列名（如'评分分组'）
        top_n : int  展示前N个编码

        Returns
        -------
        pd.DataFrame: 行为编码，列为属性值
        """
        docs = list(self.documents)
        cta = CrossTabAnalysis()
        return cta.build_matrix(docs, self.code_system, attribute, top_n=top_n)

    def build_cooccurrence_matrix(self,
                                  min_cooccurrence: int = 2) -> pd.DataFrame:
        """
        构建编码共现矩阵。

        Returns
        -------
        pd.DataFrame: 方阵，行列均为编码名
        """
        docs = list(self.documents)
        com = CooccurrenceMatrix()
        return com.build_matrix(docs, self.code_system, min_cooccurrence=min_cooccurrence)

    # ════════════════════════════════════════════════════
    # 段落浏览
    # ════════════════════════════════════════════════════

    def browse_paragraphs(self, doc_id: str = None,
                         segment_filter: str = None) -> List[Dict]:
        """
        浏览段落（可按关键词过滤）。

        Returns
        -------
        list of dict: {doc_id, doc_name, segment, index, assigned_codes}
        """
        if doc_id:
            doc = self.get_document(doc_id)
            docs = [doc] if doc else []
        else:
            docs = list(self.documents)

        results = []
        for doc in docs:
            tagger = ParagraphTagger()
            tagged = tagger.tag(doc)

            for seg_idx, seg_info in enumerate(tagged):
                if segment_filter and segment_filter.lower() not in seg_info['segment'].lower():
                    continue

                # 查找分配给该段的编码
                assigned = []
                for code_name, code in self.code_system.all_codes.items():
                    for inst in code.instances:
                        if inst['doc_id'] == doc.id and inst.get('segment', '') in seg_info['segment']:
                            assigned.append(code_name)
                            break

                results.append({
                    'doc_id': doc.id,
                    'doc_name': doc.name,
                    'segment': seg_info['segment'],
                    'index': seg_idx,
                    'assigned_codes': list(set(assigned)),
                    'rating': doc.attributes.get('score') or doc.attributes.get('rating'),
                })

        return results

    # ════════════════════════════════════════════════════
    # 编码片段导出
    # ════════════════════════════════════════════════════

    def export_coded_segments(self, output_path: str = None) -> pd.DataFrame:
        """
        导出所有编码片段为 DataFrame/Excel。

        Returns
        -------
        pd.DataFrame
        """
        exporter = CodeExporter()
        docs = list(self.documents)
        return exporter.export_to_dataframe(docs, self.code_system)

    # ════════════════════════════════════════════════════
    # 情感分析（基于文本内容，非特定行业）
    # ════════════════════════════════════════════════════

    def analyze_sentiment_all(self, text_col: str = '_text') -> pd.DataFrame:
        """
        对所有文档做情感强度分析。

        Returns
        -------
        pd.DataFrame: 包含 文档ID, 文本, 情感强度(1-5), 情感标签, 置信度
        """
        rows = []
        for doc in self.documents:
            lvl, lbl, conf = self.sentiment_analyzer.classify(doc.text)
            rows.append({
                'doc_id': doc.id,
                'doc_name': doc.name,
                'text_preview': doc.text[:100],
                'intensity_level': lvl,
                'intensity_label': lbl,
                'confidence': conf,
            })
        return pd.DataFrame(rows)

    def detect_hidden_dissatisfaction_all(self) -> pd.DataFrame:
        """
        检测所有文档中的隐性不满。

        Returns
        -------
        pd.DataFrame: 包含 文档ID, 文本, 是否隐性不满, 原因, 负面片段
        """
        rows = []
        for doc in self.documents:
            attrs = doc.attributes
            score = (attrs.get('score') or attrs.get('评分') or
                     attrs.get('rating') or attrs.get('分值') or 5)
            h = self.hidden_neg_analyzer.detect(doc.text, score=score)
            if h.get('is_hidden_neg'):
                rows.append({
                    'doc_id': doc.id,
                    'doc_name': doc.name,
                    'text': doc.text[:200],
                    'score': score,
                    'is_hidden_neg': True,
                    'reason': h.get('reason', ''),
                    'negative_fragments': '; '.join(h.get('negative_fragments', [])),
                })
        return pd.DataFrame(rows) if rows else pd.DataFrame()

    def monthly_trend_analysis(self, date_col='日期', score_col='分值') -> pd.DataFrame:
        """月度趋势分析（文档数量变化）"""
        df = self.documents.df
        if date_col not in df.columns:
            # 尝试常见列名
            for col in df.columns:
                if '日期' in col or 'date' in col.lower() or '时间' in col or 'time' in col.lower():
                    date_col = col
                    break
            else:
                return pd.DataFrame()
        if date_col not in df.columns:
            return pd.DataFrame()
        try:
            import pandas as pd
            df2 = df.copy()
            df2['_月份'] = pd.to_datetime(df2[date_col], errors='coerce').dt.to_period('M')
            result = df2.groupby('_月份').size().reset_index(name='文档数')
            result['_月份'] = result['_月份'].astype(str)
            result = result.rename(columns={'_月份': '月份'})
            return result
        except Exception:
            return pd.DataFrame()
    # ════════════════════════════════════════════════════
    # 可视化
    # ════════════════════════════════════════════════════

    def chart_code_distribution(self, title: str = '编码分布') -> Chart:
        """编码分布饼图"""
        labels = []
        sizes = []
        for code in self.code_system.root_codes:
            labels.append(code.name)
            sizes.append(len(code.instances))

        if not labels:
            return None

        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(8, 6))
        colors = [c.color for c in self.code_system.root_codes]
        ax.pie(sizes, labels=labels, autopct='%1.1f%%',
               colors=colors, startangle=90)
        ax.set_title(title)
        plt.tight_layout()
        return Chart(fig)

    def chart_cooccurrence(self, title: str = '编码共现网络') -> Chart:
        """共现网络图"""
        matrix = self.build_cooccurrence_matrix()
        if matrix.empty or len(matrix) < 2:
            return None

        import matplotlib.pyplot as plt
        import numpy as np

        fig, ax = plt.subplots(figsize=(10, 8))
        data = matrix.values.astype(float)
        np.fill_diagonal(data, 0)

        # 热力图
        im = ax.imshow(data, cmap='Blues', aspect='auto')
        ax.set_xticks(range(len(matrix.columns)))
        ax.set_yticks(range(len(matrix.index)))
        ax.set_xticklabels(matrix.columns, rotation=45, ha='right')
        ax.set_yticklabels(matrix.index)
        ax.set_title(title)
        plt.colorbar(im, ax=ax, label='共现次数')
        plt.tight_layout()
        return Chart(fig)

    # ════════════════════════════════════════════════════
    # 报告导出
    # ════════════════════════════════════════════════════

    def generate_report(self, output_path: str,
                        format: str = 'docx',
                        include_memos: bool = True,
                        include_charts: bool = True) -> str:
        """
        生成分析报告。

        Parameters
        ----------
        output_path : str  输出文件路径
        format : str  格式：'docx', 'xlsx', 'md'
        include_memos : bool  包含备忘录
        include_charts : bool  包含图表

        Returns
        -------
        str: 输出文件路径
        """
        results = {
            'project_name': self.project_name,
            'stats': self.document_stats(),
            'code_system': self.code_system.to_dict(),
            'coded_segments': self.export_coded_segments().to_dict('records') if self.code_system.all_codes else [],
            'sentiment': self.analyze_sentiment_all().to_dict('records'),
            'hidden_neg': self.detect_hidden_dissatisfaction_all().to_dict('records') if self.detect_hidden_dissatisfaction_all().columns.any() else [],
        }

        if include_memos:
            results['memos'] = self.memo_manager.to_dict()

        if format == 'docx':
            self._generate_word_report(results, output_path)
        elif format == 'xlsx':
            self._generate_excel_report(results, output_path)
        elif format == 'md':
            self._generate_markdown_report(results, output_path)

        return output_path

    def _generate_word_report(self, results: dict, output_path: str):
        from hotel_analyzer.reporter import generate_word_report
        generate_word_report({'project': results}, output_path)

    def _generate_excel_report(self, results: dict, output_path: str):
        from hotel_analyzer.reporter import generate_excel_report
        generate_excel_report({'project': results}, output_path)

    def _generate_markdown_report(self, results: dict, output_path: str):
        from hotel_analyzer.reporter import generate_markdown_report
        generate_markdown_report({'project': results}, output_path)

    # ════════════════════════════════════════════════════
    # 学术级编码增强
    # ════════════════════════════════════════════════════

    def merge_codes(self, source_names: List[str], target_name: str):
        """
        将多个源编码合并到目标编码。

        Parameters
        ----------
        source_names : List[str]
            要合并的源编码名称列表
        target_name : str
            目标编码名称（不存在则自动创建）

        Notes
        -----
        所有源编码的实例将被重新分配给目标编码，然后删除源编码。
        保留所有文档引用和片段内容。
        """
        # 确保目标编码存在
        if self.code_system._find_code_by_name(target_name) is None:
            self.create_code(target_name)

        for src_name in source_names:
            src = self.code_system._find_code_by_name(src_name)
            if src is None:
                continue
            # 将源编码的所有实例转移到目标编码
            for inst in src.instances:
                self.code_system.add_instance(target_name, inst)
            # 删除源编码
            self.code_system.remove_code(src_name)

    def rename_code(self, old_name: str, new_name: str) -> bool:
        """
        重命名编码。

        Parameters
        ----------
        old_name : str
            原编码名称
        new_name : str
            新编码名称

        Returns
        -------
        bool
            成功返回True；原编码不存在或新编码已存在返回False。
        """
        code = self.code_system._find_code_by_name(old_name)
        if code is None:
            return False
        if self.code_system._find_code_by_name(new_name) is not None:
            return False
        code.name = new_name
        return True

    def code_query(self,
                   include_codes: List[str] = None,
                   exclude_codes: List[str] = None,
                   require_all: bool = False) -> List[Dict]:
        """
        布尔编码查询引擎。

        返回符合编码组合条件的文档列表。

        Parameters
        ----------
        include_codes : List[str], optional
            包含的编码列表。require_all=True 时文档必须同时包含所有指定编码（AND），
            require_all=False 时文档只需包含其中任意一个（OR）。
        exclude_codes : List[str], optional
            排除的编码列表。包含其中任意一个编码的文档将被排除（NOT）。
        require_all : bool, default=False
            True 表示 AND 逻辑（全部包含），False 表示 OR 逻辑（任一包含）。

        Returns
        -------
        List[Dict]
            每个元素包含：
            - doc_id : str
            - doc_name : str
            - matched_codes : List[str]
            - segment_count : int
            - text_preview : str
        """
        include_codes = include_codes or []
        exclude_codes = exclude_codes or []
        exclude_set = set(exclude_codes)

        # 构建 doc_id -> set(code_name) 映射
        doc_codes: Dict[str, set] = {}
        doc_inst_count: Dict[str, int] = {}
        for code in self.code_system.all_codes.values():
            for inst in code.instances:
                did = inst.get('doc_id')
                if did:
                    doc_codes.setdefault(did, set()).add(code.name)
                    doc_inst_count[did] = doc_inst_count.get(did, 0) + 1

        results = []
        for doc in self.documents:
            codes_on_doc = doc_codes.get(doc.id, set())

            # 排除：如果文档包含任何排除编码，跳过
            if exclude_set and codes_on_doc.intersection(exclude_set):
                continue

            # 包含条件判断
            if include_codes:
                if require_all:
                    # AND: 必须包含所有指定编码
                    if not set(include_codes).issubset(codes_on_doc):
                        continue
                else:
                    # OR: 必须包含至少一个指定编码
                    if not codes_on_doc.intersection(include_codes):
                        continue

            # 匹配的编码列表
            if include_codes:
                matched = [c for c in include_codes if c in codes_on_doc]
            else:
                matched = list(codes_on_doc)

            results.append({
                'doc_id': doc.id,
                'doc_name': doc.name,
                'matched_codes': matched,
                'segment_count': doc_inst_count.get(doc.id, 0),
                'text_preview': doc.text[:200],
            })

        return results

    def get_coding_density(self) -> pd.DataFrame:
        """
        计算每个文档的编码密度。

        密度计算公式: density = total_instances / text_length * 1000

        Returns
        -------
        pd.DataFrame
            包含列: [doc_id, doc_name, total_codes, total_instances,
                     text_length, density]
        """
        rows = []
        for doc in self.documents:
            total_instances = 0
            unique_codes = set()
            for code in self.code_system.all_codes.values():
                for inst in code.instances:
                    if inst.get('doc_id') == doc.id:
                        total_instances += 1
                        unique_codes.add(code.name)
            text_len = len(doc.text)
            density = (total_instances / text_len * 1000) if text_len > 0 else 0.0
            rows.append({
                'doc_id': doc.id,
                'doc_name': doc.name,
                'total_codes': len(unique_codes),
                'total_instances': total_instances,
                'text_length': text_len,
                'density': round(density, 4),
            })
        return pd.DataFrame(rows)

    def get_uncoded_documents(self) -> List[Document]:
        """
        返回没有任何编码实例的文档列表。

        Returns
        -------
        List[Document]
            未被编码的文档列表。
        """
        coded_doc_ids = set()
        for code in self.code_system.all_codes.values():
            for inst in code.instances:
                did = inst.get('doc_id')
                if did:
                    coded_doc_ids.add(did)
        return [doc for doc in self.documents if doc.id not in coded_doc_ids]

    def export_codebook(self, output_path: str = None) -> str:
        """
        导出完整编码手册为 Markdown 格式。

        包含每个编码的名称、描述、颜色、父编码、实例数量，
        以及最多 3 个示例片段。

        Parameters
        ----------
        output_path : str, optional
            如果提供，将 Markdown 写入该文件。

        Returns
        -------
        str
            Markdown 格式的编码手册字符串。
        """
        from datetime import datetime
        lines = []
        lines.append('# 编码手册 (Codebook)\n\n')
        lines.append(f'生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n')
        lines.append('---\n\n')

        for code in self.code_system.all_codes.values():
            lines.append(f'## {code.name}\n\n')
            lines.append(f'- **描述**: {code.description or "无"}\n')
            lines.append(f'- **颜色**: {code.color}\n')
            parent_name = code.parent.name if code.parent else '(无)'
            lines.append(f'- **父编码**: {parent_name}\n')
            n_inst = len(code.instances)
            lines.append(f'- **编码实例数**: {n_inst}\n\n')

            if code.instances:
                lines.append('### 示例片段\n\n')
                for i, inst in enumerate(code.instances[:3]):
                    seg = inst.get('segment', '')
                    lines.append(f'{i+1}. "{seg}"\n')
                if n_inst > 3:
                    lines.append(f'\n*...以及 {n_inst - 3} 个其他片段*\n')
            lines.append('\n---\n\n')

        content = ''.join(lines)
        if output_path:
            Path(output_path).write_text(content, encoding='utf-8')
        return content

    def get_audit_trail(self) -> pd.DataFrame:
        """
        构建编码审计追踪。

        由于当前版本不存储时间戳，使用当前时间作为近似值。

        Returns
        -------
        pd.DataFrame
            包含列: [timestamp, action, code, doc_id, segment_preview]
        """
        rows = []
        now = pd.Timestamp.now()
        for code in self.code_system.all_codes.values():
            for inst in code.instances:
                seg_preview = inst.get('segment', '')[:100]
                rows.append({
                    'timestamp': now,
                    'action': 'assign_code',
                    'code': code.name,
                    'doc_id': inst.get('doc_id', ''),
                    'segment_preview': seg_preview,
                })
        return pd.DataFrame(rows)

    def load_intercoder_codes(self, csv_path: str,
                             doc_id_col: str = 'doc_id',
                             code_col: str = 'code_name') -> Dict[str, List[str]]:
        """
        从 CSV 文件加载第二位编码者的编码结果，用于计算 Cohen's Kappa。

        CSV 格式要求：至少包含 ``doc_id`` 列和 ``code_name`` 列，
        每行代表一个编码实例（即一个 doc_id 可以出现多行）。

        Parameters
        ----------
        csv_path : str  CSV 文件路径
        doc_id_col : str  文档ID列名，默认 'doc_id'
        code_col : str  编码名列名，默认 'code_name'

        Returns
        -------
        Dict[str, List[str]]: {doc_id: [code_name, ...]}
        """
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        if doc_id_col not in df.columns or code_col not in df.columns:
            raise ValueError(
                f"CSV 必须包含 '{doc_id_col}' 和 '{code_col}' 列。"
                f"当前列: {list(df.columns)}")
        result: Dict[str, List[str]] = {}
        for _, row in df.iterrows():
            did = str(row[doc_id_col])
            cname = str(row[code_col])
            result.setdefault(did, []).append(cname)
        return result

    def get_coding_saturation(self) -> pd.DataFrame:
        """
        计算编码饱和度：每个文档贡献了多少个新编码。

        按文档加载顺序（collection 中顺序）依次计算，
        记录到该文档为止累计出现了多少新编码。
        新编码越多，说明该文档带来越多新信息。

        Returns
        -------
        pd.DataFrame: [doc_id, doc_name, total_codes, new_codes, cumulative_codes, saturation_rate]
        """
        rows = []
        seen: set = set()
        for doc in self.documents:
            on_doc = {
                c.name for c in self.code_system.all_codes.values()
                if any(i.get('doc_id') == doc.id for i in c.instances)
            }
            new = len(on_doc - seen)
            seen |= on_doc
            total = len(on_doc)
            rows.append({
                'doc_id': doc.id,
                'doc_name': doc.name,
                'total_codes': total,
                'new_codes': new,
                'cumulative_codes': len(seen),
                'saturation_rate': round(new / max(total, 1), 3),
            })
        return pd.DataFrame(rows)

    def intercoder_reliability(self, other_codes: dict) -> dict:
        """
        计算本编码者与另一位编码者之间的 Cohen's Kappa 系数。

        Parameters
        ----------
        other_codes : dict
            另一位编码者的编码结果，格式为 {doc_id: [code_names]}

        Returns
        -------
        dict
            {
                'kappa': float,           # Cohen's Kappa 系数
                'agreement_pct': float,    # 一致百分比
                'n_docs': int,            # 共同文档数
                'n_codes': int,           # 共同编码数
                'confusion_table': dict,   # {tp, tn, fp, fn}
            }
        """
        # 构建本编码者的编码映射 {doc_id: {code_names}}
        my_codes: Dict[str, set] = {}
        for code in self.code_system.all_codes.values():
            for inst in code.instances:
                did = inst.get('doc_id')
                if did:
                    my_codes.setdefault(did, set()).add(code.name)

        # 所有唯一的编码名
        all_code_names = set()
        for codes in my_codes.values():
            all_code_names.update(codes)
        for codes in other_codes.values():
            all_code_names.update(codes)
        all_code_names = sorted(all_code_names)

        # 所有唯一的文档ID
        all_doc_ids = sorted(set(my_codes.keys()) | set(other_codes.keys()))
        n_docs = len(all_doc_ids)
        n_codes = len(all_code_names)

        if n_docs == 0 or n_codes == 0:
            return {
                'kappa': 0.0,
                'agreement_pct': 0.0,
                'n_docs': n_docs,
                'n_codes': n_codes,
                'confusion_table': {},
            }

        # 为每个 (doc_id, code_name) 对计算编码一致性
        n_total = n_docs * n_codes
        tp = fp = fn = tn = 0

        for did in all_doc_ids:
            my_set = my_codes.get(did, set())
            other_set = set(other_codes.get(did, []))
            for cname in all_code_names:
                my_val = 1 if cname in my_set else 0
                other_val = 1 if cname in other_set else 0
                if my_val == 1 and other_val == 1:
                    tp += 1
                elif my_val == 0 and other_val == 0:
                    tn += 1
                elif my_val == 1 and other_val == 0:
                    fp += 1
                else:
                    fn += 1

        p_o = (tp + tn) / n_total if n_total > 0 else 0.0

        # 计算预期一致率 (Cohen's Kappa)
        n_this = tp + fp
        n_other = tp + fn
        p_this = n_this / n_total
        p_other = n_other / n_total
        p_not_this = (tn + fn) / n_total
        p_not_other = (tn + fp) / n_total
        p_e = p_this * p_other + p_not_this * p_not_other

        kappa = (p_o - p_e) / (1 - p_e) if (1 - p_e) > 0 else 0.0

        return {
            'kappa': round(kappa, 4),
            'agreement_pct': round(p_o * 100, 2),
            'n_docs': n_docs,
            'n_codes': n_codes,
            'confusion_table': {
                'true_positive': tp,
                'true_negative': tn,
                'false_positive': fp,
                'false_negative': fn,
            },
        }

    # ════════════════════════════════════════════════════
    # 项目保存/加载
    # ════════════════════════════════════════════════════

    def save_project(self, project_path: str):
        """
        保存完整项目为 .qda 文件（ZIP格式）。

        .qda 文件包含：
        - documents.json    文档集合
        - codes.json        编码系统
        - memos.json        备忘录
        - project.json      项目配置
        """
        path = Path(project_path)
        if path.suffix != '.qda':
            path = Path(str(path) + '.qda')

        self.project_path = str(path)

        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # 文档
            docs_data = [d.to_dict() for d in self.documents]
            zf.writestr('documents.json', json.dumps(docs_data, ensure_ascii=False, indent=2))

            # 编码系统
            zf.writestr('codes.json', json.dumps(self.code_system.to_dict(), ensure_ascii=False, indent=2))

            # 备忘录
            zf.writestr('memos.json', json.dumps(self.memo_manager.to_dict(), ensure_ascii=False, indent=2))

            # 项目配置
            project_cfg = {
                'name': self.project_name,
                'version': self.VERSION,
                'saved_at': pd.Timestamp.now().isoformat(),
            }
            zf.writestr('project.json', json.dumps(project_cfg, ensure_ascii=False, indent=2))

        return str(path)

    @classmethod
    def load_project(cls, project_path: str) -> 'QDAApplication':
        """加载 .qda 项目文件"""
        path = Path(project_path)
        app = cls(project_name='加载项目')

        with zipfile.ZipFile(path, 'r') as zf:
            # 文档
            with zf.open('documents.json') as f:
                docs_data = json.load(f)
            docs = [Document.from_dict(d) for d in docs_data]
            app.documents = DocumentCollection(docs)

            # 编码系统
            with zf.open('codes.json') as f:
                codes_data = json.load(f)
            app.code_system = CodeSystem.from_dict(codes_data)

            # 备忘录
            with zf.open('memos.json') as f:
                memos_data = json.load(f)
            app.memo_manager = MemoManager.from_dict(memos_data)

            # 项目配置
            with zf.open('project.json') as f:
                cfg = json.load(f)
            app.project_name = cfg.get('name', '未命名项目')
            app.project_path = str(path)

        return app

    # ════════════════════════════════════════════════════
    # 摘要
    # ════════════════════════════════════════════════════

    def summary(self) -> dict:
        """返回项目摘要"""
        return {
            'project_name': self.project_name,
            'total_documents': len(self.documents),
            'total_codes': len(self.code_system.all_codes),
            'total_memos': self.memo_manager.summary()['总备忘录数'],
            'code_instances': sum(len(c.instances) for c in self.code_system.all_codes.values()),
            'attributes': list(self.documents.df.columns),
        }

    def __repr__(self):
        return (f"<QDAApplication '{self.project_name}': "
                f"{len(self.documents)}文档, "
                f"{len(self.code_system.all_codes)}编码, "
                f"{self.memo_manager.summary()['总备忘录数']}备忘录>")
