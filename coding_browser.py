"""
编码与检索模块 - QDA核心
========================
提供：层级代码树、段落标注、交叉分析、共现矩阵、全文检索、片段导出

本模块不包含任何行业特定内容，完全基于通用文本处理。
"""

import re
import pandas as pd
import numpy as np
from collections import defaultdict
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

# ══════════════════════════════════════════════════════════════════
# 代码定义
# ══════════════════════════════════════════════════════════════════

class Code:
    """
    单个代码（Code）。

    grounded_stage: 扎根理论阶段，可选值：
      'open' / '开放编码'      - 开放编码阶段
      'axial' / '轴心编码'     - 轴心编码阶段
      'selective' / '选择性编码' - 选择性编码阶段
    """
    def __init__(self, name: str, color: str = None,
                 parent: 'Code' = None, description: str = '',
                 grounded_stage: str = None):
        self.id = id(self)
        self.name = name
        self.color = color or '#1a73e8'
        self.parent = parent
        self.description = description
        self.grounded_stage = grounded_stage  # 扎根理论阶段
        self.children: List['Code'] = []
        self.instances: List[Dict] = []   # [{doc_id, text, start, end, memo}]
        self.memos: List[Dict] = []       # [{text, type}]
        self.created_at = None

    def __repr__(self):
        return f'Code({self.name}, {len(self.instances)}次)'


class CodeSystem:
    """
    编码系统。
    支持层级代码、自由编码、实例管理、序列化。
    """

    COLOR_PALETTE = [
        '#e53935', '#d81b60', '#8e24aa', '#5e35b1',
        '#3949ab', '#1e88e5', '#039be5', '#00acc1',
        '#00897b', '#43a047', '#7cb342', '#c0ca33',
        '#fdd835', '#ffb300', '#fb8c00', '#f4511e',
        '#6d4c41', '#546e7a', '#757575', '#37474f',
    ]

    def __init__(self, name: str = '新编码系统'):
        self.name = name
        self.root_codes: List[Code] = []
        self.all_codes: Dict[int, Code] = {}  # {code_id: Code}
        self.next_color_idx = 0

    def next_color(self) -> str:
        c = self.COLOR_PALETTE[self.next_color_idx % len(self.COLOR_PALETTE)]
        self.next_color_idx += 1
        return c

    # ---- 增删 ----
    def add_code(self, name: str, color: str = None,
                description: str = '', parent_name: str = None) -> Code:
        """添加代码（支持层级）"""
        if parent_name:
            parent = self._find_code_by_name(parent_name)
            if not parent:
                return self.add_code(name, color, description, None)
            code = Code(name, color or self.next_color(), parent=parent,
                       description=description)
            parent.children.append(code)
        else:
            code = Code(name, color or self.next_color(),
                       description=description)
            self.root_codes.append(code)
        self.all_codes[code.id] = code
        return code

    def _find_code_by_name(self, name: str) -> Optional[Code]:
        for c in self.all_codes.values():
            if c.name == name:
                return c
        return None

    def remove_code(self, code_name: str) -> bool:
        code = self._find_code_by_name(code_name)
        if not code:
            return False
        if code.parent:
            code.parent.children.remove(code)
        else:
            self.root_codes.remove(code)
        del self.all_codes[code.id]
        return True

    # ---- 实例管理 ----
    def add_instance(self, code_name: str, instance: Dict) -> Dict:
        """添加编码实例"""
        code = self._find_code_by_name(code_name)
        if not code:
            code = self.add_code(code_name)
        instance = {**instance}
        code.instances.append(instance)
        return instance

    def get_instances(self, code_name: str = None,
                    doc_id: str = None) -> List[Dict]:
        results = []
        codes = ([self._find_code_by_name(code_name)] if code_name
                 else list(self.all_codes.values()))
        for code in codes:
            if not code:
                continue
            for inst in code.instances:
                if doc_id is None or inst.get('doc_id') == doc_id:
                    results.append({**inst, 'code_name': code.name,
                                  'code_color': code.color})
        return results

    # ---- 序列化 ----
    def to_dict(self) -> dict:
        def code_to_dict(code: Code, parent_id=None) -> dict:
            return {
                'id': code.id,
                'name': code.name,
                'color': code.color,
                'parent_id': parent_id,
                'description': code.description,
                'grounded_stage': code.grounded_stage,
                'children': [code_to_dict(c, code.id) for c in code.children],
                'instances': code.instances,
                'memos': code.memos,
            }
        return {
            'name': self.name,
            'codes': [code_to_dict(c) for c in self.root_codes],
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'CodeSystem':
        cs = cls(data.get('name', '加载的编码系统'))

        def dict_to_code(d: dict, parent: Code = None) -> Code:
            code = Code(
                name=d['name'],
                color=d.get('color'),
                parent=parent,
                description=d.get('description', ''),
                grounded_stage=d.get('grounded_stage'),
            )
            code.id = d.get('id', id(code))
            code.instances = d.get('instances', [])
            code.memos = d.get('memos', [])
            cs.all_codes[code.id] = code
            if parent:
                parent.children.append(code)
            else:
                cs.root_codes.append(code)
            for child_d in d.get('children', []):
                dict_to_code(child_d, code)
            return code

        for cd in data.get('codes', []):
            dict_to_code(cd)
        return cs

    def summary(self) -> dict:
        return {
            '代码总数': len(self.all_codes),
            '顶级代码': len(self.root_codes),
            '总实例数': sum(len(c.instances) for c in self.all_codes.values()),
        }


# ══════════════════════════════════════════════════════════════════
# 段落标注器

class ParagraphTagger:
    """
    将文档切分为段落/句子单元，供用户逐段编码。
    切分策略：中文句号(。！？)分段，英文句子(.!?)分段。
    """

    def __init__(self, min_segment_len: int = 5, max_segment_len: int = 500):
        self.min_segment_len = min_segment_len
        self.max_segment_len = max_segment_len

    def tag(self, document: 'Document') -> List[Dict]:
        """
        将文档切分为段落列表。

        Returns:
            list of dict: [{
                'segment': str,          # 段落文本
                'start': int,            # 在原文中起始位置
                'end': int,              # 结束位置
                'index': int,            # 段落序号
            }]
        """
        if not hasattr(document, 'text'):
            return []

        text = document.text
        # 按句子切分
        # 保留分隔符
        pattern = r'(?<=[。！？.!?])'
        parts = re.split(pattern, text)
        # 合并过短的段落
        segments = []
        current = ''
        current_start = 0

        for part in parts:
            if not part.strip():
                continue
            if len(current) + len(part) <= self.max_segment_len:
                if not current:
                    current_start = text.find(part)
                current += part
            else:
                if current.strip():
                    segments.append({
                        'segment': current.strip(),
                        'start': current_start,
                        'end': current_start + len(current),
                        'index': len(segments),
                    })
                current = part
                current_start = text.find(part)

        if current.strip():
            segments.append({
                'segment': current.strip(),
                'start': current_start,
                'end': current_start + len(current),
                'index': len(segments),
            })

        # 过滤过短段落
        segments = [s for s in segments
                   if len(s['segment']) >= self.min_segment_len]

        # 重新编号
        for i, s in enumerate(segments):
            s['index'] = i

        return segments


# ══════════════════════════════════════════════════════════════════
# 交叉分析
# ══════════════════════════════════════════════════════════════════

class CrossTabAnalysis:
    """
    交叉分析表：代码 × 属性（分组）统计。
    """

    def build_matrix(self,
                    documents: List['Document'],
                    code_system: 'CodeSystem',
                    attribute: str,
                    top_n: int = 10) -> pd.DataFrame:
        """
        构建 编码 × 属性 的交叉表。

        Parameters
        ----------
        documents : list of Document
        code_system : CodeSystem
        attribute : str  文档属性名（如'评分分组'、'出游类型'）
        top_n : int  只展示编码实例数前N的代码

        Returns
        -------
        pd.DataFrame: 行=代码名，列=属性值，值=片段数
        """
        if not documents:
            return pd.DataFrame()

        # 收集属性值
        attr_values: set = set()
        doc_attrs: Dict[str, str] = {}
        for doc in documents:
            val = doc.attributes.get(attribute, '未分类')
            attr_values.add(val)
            doc_attrs[doc.id] = val

        # 统计每个(代码, 属性值)的实例数
        matrix: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))

        for code in code_system.all_codes.values():
            code_name = code.name
            for inst in code.instances:
                doc_id = inst['doc_id']
                attr_val = doc_attrs.get(doc_id, '未分类')
                matrix[code_name][attr_val] += 1

        # 按总实例数排序，取TopN
        code_totals = {c: sum(v.values()) for c, v in matrix.items()}
        top_codes = sorted(code_totals, key=lambda x: -code_totals[x])[:top_n]

        attr_list = sorted(attr_values)
        rows = []
        for code_name in top_codes:
            row = {'代码': code_name}
            for attr_val in attr_list:
                row[attr_val] = matrix[code_name][attr_val]
            row['总计'] = sum(matrix[code_name].values())
            rows.append(row)

        return pd.DataFrame(rows).set_index('代码')


# ══════════════════════════════════════════════════════════════════
# 共现矩阵
# ══════════════════════════════════════════════════════════════════

class CooccurrenceMatrix:
    """
    构建代码共现矩阵。
    两个代码在同一文档的片段中出现，即为共现。
    """

    def build_matrix(self,
                    documents: List['Document'],
                    code_system: 'CodeSystem',
                    min_cooccurrence: int = 1) -> pd.DataFrame:
        """
        Returns
        -------
        pd.DataFrame: 方阵，行列均为代码名
        """
        # 收集每个文档中的代码集合
        doc_codes: Dict[str, set] = defaultdict(set)
        for code in code_system.all_codes.values():
            for inst in code.instances:
                doc_codes[inst['doc_id']].add(code.name)

        # 统计共现
        code_names = sorted(code_system.all_codes.keys(),
                          key=lambda x: code_system.all_codes[x].name)
        code_names_display = [code_system.all_codes[cid].name for cid in code_names]

        co_matrix = pd.DataFrame(
            np.zeros((len(code_names), len(code_names)), dtype=int),
            index=code_names_display,
            columns=code_names_display,
        )

        for doc_id, codes_in_doc in doc_codes.items():
            codes_list = sorted(codes_in_doc)
            for i, c1 in enumerate(codes_list):
                for c2 in codes_list[i+1:]:
                    co_matrix.loc[c1, c2] += 1
                    co_matrix.loc[c2, c1] += 1

        # 对角线：每个代码的实例数
        for cid, cname in zip(code_names, code_names_display):
            co_matrix.loc[cname, cname] = len(code_system.all_codes[cid].instances)

        return co_matrix


# ══════════════════════════════════════════════════════════════════
# 全文检索
# ══════════════════════════════════════════════════════════════════

class AdvancedSearch:
    """
    高级检索引擎。
    支持正则表达式、上下文窗口高亮。
    """

    def search_documents(self,
                        documents: List['Document'],
                        pattern: str,
                        use_regex: bool = True,
                        context_chars: int = 50) -> List[Dict]:
        """
        搜索所有文档。

        Parameters
        ----------
        documents : list of Document
        pattern : str  关键词或正则
        use_regex : bool  是否将pattern作为正则解析
        context_chars : int  匹配点前后显示字符数

        Returns
        -------
        list of dict: [{
            'doc_id': str,
            'doc_name': str,
            'match': str,       # 匹配的文本
            'position': int,    # 在原文中的位置
            'context_before': str,
            'context_after': str,
            'context': str,     # 包含前后的完整上下文
        }]
        """
        results = []

        for doc in documents:
            text = doc.text
            if use_regex:
                try:
                    flags = re.IGNORECASE
                    regex = re.compile(pattern, flags)
                except re.error:
                    regex = re.compile(re.escape(pattern), flags)
            else:
                regex = re.compile(re.escape(pattern), re.IGNORECASE)

            for m in regex.finditer(text):
                start = m.start()
                end = m.end()
                ctx_before = text[max(0, start - context_chars):start]
                ctx_after = text[end:min(len(text), end + context_chars)]
                results.append({
                    'doc_id': doc.id,
                    'doc_name': doc.name,
                    'match': m.group(),
                    'position': start,
                    'context_before': ctx_before,
                    'context_after': ctx_after,
                    'context': f'...{ctx_before}[{m.group()}]{ctx_after}...',
                })

        return results

    def search_codes(self,
                    code_system: 'CodeSystem',
                    pattern: str,
                    use_regex: bool = False) -> List[Dict]:
        """
        在代码名称和描述中搜索。

        Returns
        -------
        list of dict: [{'code': Code, 'match': str}]
        """
        results = []
        for code in code_system.all_codes.values():
            if use_regex:
                try:
                    if re.search(pattern, code.name, re.IGNORECASE):
                        results.append({'code': code, 'match': code.name})
                except re.error:
                    pass
            else:
                if pattern.lower() in code.name.lower():
                    results.append({'code': code, 'match': code.name})
        return results


# ══════════════════════════════════════════════════════════════════
# 段落浏览器
# ══════════════════════════════════════════════════════════════════

class SegmentBrowser:
    """
    交互式段落浏览与编码。
    提供分段视图、代码分配、片段导出功能。
    """

    def __init__(self, code_system: 'CodeSystem' = None):
        self.code_system = code_system

    def browse(self,
              document: 'Document',
              tagged_segments: List[Dict] = None) -> List[Dict]:
        """
        浏览文档的所有段落及已分配的代码。

        Returns
        -------
        list of dict: [{
            'segment': str, 'index': int,
            'assigned_codes': [code_name, ...],
            'start': int, 'end': int,
        }]
        """
        if tagged_segments is None:
            tagger = ParagraphTagger()
            tagged_segments = tagger.tag(document)

        results = []
        for seg in tagged_segments:
            # 查找分配给该段的代码
            assigned = []
            if self.code_system:
                for code_name, code in self.code_system.all_codes.items():
                    for inst in code.instances:
                        if inst['doc_id'] == document.id:
                            seg_text = seg['segment']
                            inst_text = inst.get('text', '')
                            # 如果实例文本在该段中
                            if inst_text and inst_text in seg_text:
                                assigned.append(code.name)
            results.append({
                'segment': seg['segment'],
                'index': seg['index'],
                'start': seg['start'],
                'end': seg['end'],
                'assigned_codes': list(set(assigned)),
            })
        return results

    def get_segment(self,
                    document: 'Document',
                    segment_index: int,
                    window_chars: int = 100) -> Dict:
        """
        获取特定段落的详细信息。
        """
        tagger = ParagraphTagger()
        segments = tagger.tag(document)

        if segment_index < 0 or segment_index >= len(segments):
            return {}

        seg = segments[segment_index]
        # 收集该段关联的编码实例
        instances = []
        if self.code_system:
            for code_name, code in self.code_system.all_codes.items():
                for inst in code.instances:
                    if inst['doc_id'] == document.id:
                        inst_text = inst.get('text', '')
                        if inst_text and inst_text in seg['segment']:
                            instances.append({
                                'code_name': code.name,
                                'code_color': code.color,
                                'text': inst_text,
                                'memo': inst.get('memo', ''),
                            })

        return {
            'segment': seg['segment'],
            'index': segment_index,
            'start': seg['start'],
            'end': seg['end'],
            'instances': instances,
        }


# ══════════════════════════════════════════════════════════════════
# 编码片段导出器
# ══════════════════════════════════════════════════════════════════

class CodeExporter:
    """
    将编码结果导出为多种格式（DataFrame、Excel、CSV）。
    """

    def export_to_dataframe(self,
                           documents: List['Document'],
                           code_system: 'CodeSystem',
                           text_col: str = '_text') -> pd.DataFrame:
        """
        导出所有编码片段为 DataFrame。

        Returns
        -------
        pd.DataFrame: 列 = doc_id, doc_name, code, code_color,
                      segment, start, end, memo, [attributes...]
        """
        rows = []
        doc_map = {d.id: d for d in documents}

        for code in code_system.all_codes.values():
            for inst in code.instances:
                doc = doc_map.get(inst['doc_id'])
                row = {
                    'doc_id': inst['doc_id'],
                    'doc_name': doc.name if doc else inst['doc_id'],
                    'code': code.name,
                    'code_color': code.color,
                    'segment': inst.get('text', '') or inst.get('segment', ''),
                    'start': inst.get('start', -1),
                    'end': inst.get('end', -1),
                    'memo': inst.get('memo', ''),
                }
                # 添加文档属性
                if doc:
                    for k, v in doc.attributes.items():
                        row[k] = v
                rows.append(row)

        if not rows:
            return pd.DataFrame()

        df = pd.DataFrame(rows)
        return df

    def export_to_excel(self,
                       documents: List['Document'],
                       code_system: 'CodeSystem',
                       output_path: str):
        """导出到Excel（多个Sheet）"""
        df = self.export_to_dataframe(documents, code_system)

        if df.empty:
            return

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Sheet1: 全部片段
            df.to_excel(writer, sheet_name='编码片段', index=False)

            # Sheet2: 代码统计
            code_stats = df.groupby('code').agg(
                片段数=('segment', 'count'),
            ).reset_index().sort_values('片段数', ascending=False)
            code_stats.to_excel(writer, sheet_name='代码统计', index=False)

            # Sheet3: 代码×文档矩阵
            if 'doc_name' in df.columns:
                pivot = pd.crosstab(df['code'], df['doc_name'])
                pivot.to_excel(writer, sheet_name='代码×文档矩阵')

    def export_code_book(self,
                        code_system: 'CodeSystem',
                        output_path: str = None) -> pd.DataFrame:
        """
        导出编码本（代码清单）。

        Returns
        -------
        pd.DataFrame: code, color, parent, instances, memos
        """
        rows = []

        def collect(code, parent_name=''):
            rows.append({
                '代码': code.name,
                '颜色': code.color,
                '父级': parent_name,
                '片段数': len(code.instances),
                '描述': code.description,
                '备忘录': '; '.join(m.get('text', '') for m in code.memos),
            })
            for child in code.children:
                collect(child, code.name)

        for root in code_system.root_codes:
            collect(root)

        df = pd.DataFrame(rows)
        if output_path:
            df.to_excel(output_path, index=False)
        return df
