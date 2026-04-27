"""
备忘录模块 - 模拟MAXQDA备忘录系统
支持：文档级、代码级、段落级备忘录
"""

import pandas as pd
from datetime import datetime
from pathlib import Path

# ==================== 备忘录条目 ====================

class Memo:
    """单条备忘录"""
    def __init__(self, text, memo_type='general', author='研究用户'):
        self.id = id(self)
        self.text = text
        self.type = memo_type  # 'general' | 'method' | 'theory' | 'reflection'
        self.author = author
        self.created_at = datetime.now()
        self.updated_at = datetime.now()
        self.tags = []

    def __repr__(self):
        return f'Memo({self.type}, {self.created_at.strftime("%Y-%m-%d")})'


# ==================== 备忘录管理器 ====================

class MemoManager:
    """
    备忘录管理器，模拟MAXQDA的备忘录系统：
    - 支持3种备忘录：文档级、代码级、段落级
    - 层级结构，支持搜索和标签
    - 导出为Markdown/Word
    """

    MEMO_TYPES = {
        'general':   {'label': '📝 一般备注',   'color': '#888888'},
        'method':    {'label': '🔬 方法笔记',   'color': '#1565c0'},
        'theory':    {'label': '💡 理论笔记',   'color': '#6a1b9a'},
        'reflection':{'label': '🤔 反思笔记',   'color': '#e65100'},
        'finding':   {'label': '🔍 研究发现',   'color': '#2e7d32'},
        'question':  {'label': '❓ 研究问题',    'color': '#c62828'},
    }

    def __init__(self):
        self.doc_memos = {}    # {doc_id: [Memo, ...]}
        self.code_memos = {}   # {code_id: [Memo, ...]}
        self.segment_memos = {} # {(doc_id, text, start): [Memo, ...]}

        # 全局备忘录（项目级）
        self.project_memos = []

    # ---- 创建备忘录 ----
    def add_doc_memo(self, doc_id, text, memo_type='general', author='研究用户'):
        """添加文档级备忘录"""
        if doc_id not in self.doc_memos:
            self.doc_memos[doc_id] = []
        memo = Memo(text, memo_type, author)
        self.doc_memos[doc_id].append(memo)
        return memo

    def add_code_memo(self, code_id, text, memo_type='general', author='研究用户'):
        """添加代码级备忘录"""
        if code_id not in self.code_memos:
            self.code_memos[code_id] = []
        memo = Memo(text, memo_type, author)
        self.code_memos[code_id].append(memo)
        return memo

    def add_segment_memo(self, doc_id, text, start, text_content, memo_type='general', author='研究用户'):
        """添加段落级备忘录"""
        key = (doc_id, text[:50] if text else '', start or 0)
        if key not in self.segment_memos:
            self.segment_memos[key] = []
        memo = Memo(text_content, memo_type, author)
        self.segment_memos[key].append(memo)
        return memo

    def add_project_memo(self, text, memo_type='general', author='研究用户'):
        """添加项目级备忘录"""
        memo = Memo(text, memo_type, author)
        self.project_memos.append(memo)
        return memo

    # ---- 读取备忘录 ----
    def get_doc_memos(self, doc_id):
        return self.doc_memos.get(doc_id, [])

    def get_code_memos(self, code_id):
        return self.code_memos.get(code_id, [])

    def get_segment_memos(self, doc_id, text='', start=None):
        key = (doc_id, text[:50] if text else '', start or 0)
        return self.segment_memos.get(key, [])

    def get_project_memos(self):
        return self.project_memos

    # ---- 修改/删除 ----
    def update_memo(self, memo, new_text):
        memo.text = new_text
        memo.updated_at = datetime.now()

    def delete_memo(self, memo, target_dict):
        """从目标字典中删除备忘录"""
        for key, memos in list(target_dict.items()):
            if memo in memos:
                memos.remove(memo)
                if not memos:
                    del target_dict[key]
                return True
        return False

    def delete_doc_memo(self, doc_id, memo):
        return self.delete_memo(memo, self.doc_memos)

    def delete_code_memo(self, code_id, memo):
        return self.delete_memo(memo, self.code_memos)

    # ---- 搜索 ----
    def search_memos(self, keyword, scope='all'):
        """
        搜索备忘录

        Args:
            keyword: 关键词
            scope: 'all' | 'doc' | 'code' | 'segment' | 'project'

        Returns:
            list: [(memo, context_info)]
        """
        results = []
        keyword = keyword.lower()

        targets = []
        if scope in ('all', 'doc'):
            for doc_id, memos in self.doc_memos.items():
                targets.append((doc_id, 'doc', memos))
        if scope in ('all', 'code'):
            for code_id, memos in self.code_memos.items():
                targets.append((code_id, 'code', memos))
        if scope in ('all', 'segment'):
            for key, memos in self.segment_memos.items():
                targets.append((key, 'segment', memos))
        if scope in ('all', 'project'):
            targets.append((None, 'project', self.project_memos))

        for identifier, id_type, memos in targets:
            for memo in memos:
                if keyword in memo.text.lower():
                    ctx = self._get_context_label(identifier, id_type)
                    results.append((memo, ctx))

        return results

    def _get_context_label(self, identifier, id_type):
        if id_type == 'doc':
            return f'文档: {identifier}'
        elif id_type == 'code':
            return f'代码ID: {identifier}'
        elif id_type == 'segment':
            doc_id, text, start = identifier
            return f'段落: doc={doc_id}, start={start}'
        elif id_type == 'project':
            return '项目级'

    # ---- 标签系统 ----
    def tag_memo(self, memo, tag):
        if tag not in memo.tags:
            memo.tags.append(tag)

    def untag_memo(self, memo, tag):
        if tag in memo.tags:
            memo.tags.remove(tag)

    def get_memos_by_tag(self, tag):
        """获取含特定标签的所有备忘录"""
        results = []
        for id_type, memo_dict in [('doc', self.doc_memos),
                                     ('code', self.code_memos),
                                     ('segment', self.segment_memos)]:
            for identifier, memos in memo_dict.items():
                for memo in memos:
                    if tag in memo.tags:
                        results.append((memo, identifier, id_type))
        for memo in self.project_memos:
            if tag in memo.tags:
                results.append((memo, None, 'project'))
        return results

    def all_tags(self):
        """获取所有已使用的标签"""
        tags = set()
        for id_type, memo_dict in [('doc', self.doc_memos),
                                     ('code', self.code_memos),
                                     ('segment', self.segment_memos)]:
            for memos in memo_dict.values():
                for memo in memos:
                    tags.update(memo.tags)
        for memo in self.project_memos:
            tags.update(memo.tags)
        return sorted(tags)

    # ---- 导出 ----
    def export_to_markdown(self, filepath=None):
        """
        导出所有备忘录为Markdown

        Returns:
            str: Markdown文本
        """
        lines = ['# 质性分析 - 备忘录\n']
        lines.append(f'生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
        lines.append('---\n\n')

        # 项目级
        if self.project_memos:
            lines.append('## 项目级备忘录\n\n')
            for memo in self.project_memos:
                lines.append(self._memo_to_md(memo, '项目级'))
            lines.append('\n')

        # 文档级
        if self.doc_memos:
            lines.append('## 文档级备忘录\n\n')
            for doc_id, memos in self.doc_memos.items():
                lines.append(f'### 文档 {doc_id}\n\n')
                for memo in memos:
                    lines.append(self._memo_to_md(memo))
                lines.append('\n')

        # 代码级
        if self.code_memos:
            lines.append('## 代码级备忘录\n\n')
            for code_id, memos in self.code_memos.items():
                lines.append(f'### 代码 {code_id}\n\n')
                for memo in memos:
                    lines.append(self._memo_to_md(memo))
                lines.append('\n')

        content = ''.join(lines)
        if filepath:
            Path(filepath).write_text(content, encoding='utf-8')
        return content

    def _memo_to_md(self, memo, extra_label=''):
        type_info = self.MEMO_TYPES.get(memo.type, {})
        label = type_info.get('label', memo.type)
        lines = [
            f'- **{label}**',
            f'  - 作者: {memo.author}',
            f'  - 创建: {memo.created_at.strftime("%Y-%m-%d %H:%M")}',
            f'  - 更新: {memo.updated_at.strftime("%Y-%m-%d %H:%M")}',
            f'  - 内容: {memo.text}',
        ]
        if memo.tags:
            lines.append(f'  - 标签: {" ".join(f"`{t}`" for t in memo.tags)}')
        lines.append('')
        return '\n'.join(lines)

    def export_to_dataframe(self):
        """
        导出所有备忘录为DataFrame

        Returns:
            pd.DataFrame
        """
        rows = []
        for id_type, memo_dict in [('doc', self.doc_memos),
                                     ('code', self.code_memos),
                                     ('segment', self.segment_memos)]:
            for identifier, memos in memo_dict.items():
                for memo in memos:
                    rows.append({
                        '级别': id_type,
                        '对象ID': str(identifier),
                        '类型': memo.type,
                        '作者': memo.author,
                        '创建时间': memo.created_at,
                        '更新时间': memo.updated_at,
                        '标签': ', '.join(memo.tags),
                        '内容': memo.text,
                    })
        for memo in self.project_memos:
            rows.append({
                '级别': 'project',
                '对象ID': '',
                '类型': memo.type,
                '作者': memo.author,
                '创建时间': memo.created_at,
                '更新时间': memo.updated_at,
                '标签': ', '.join(memo.tags),
                '内容': memo.text,
            })
        return pd.DataFrame(rows)

    # ---- 序列化 ----
    def to_dict(self):
        def memo_to_dict(m):
            return {
                'text': m.text,
                'type': m.type,
                'author': m.author,
                'created_at': m.created_at.isoformat(),
                'updated_at': m.updated_at.isoformat(),
                'tags': m.tags,
            }

        return {
            'doc_memos': {str(k): [memo_to_dict(m) for m in memos]
                          for k, memos in self.doc_memos.items()},
            'code_memos': {str(k): [memo_to_dict(m) for m in memos]
                           for k, memos in self.code_memos.items()},
            'segment_memos': {str(k): [memo_to_dict(m) for m in memos]
                              for k, memos in self.segment_memos.items()},
            'project_memos': [memo_to_dict(m) for m in self.project_memos],
        }

    @classmethod
    def from_dict(cls, data):
        mm = cls()

        def dict_to_memo(d):
            m = Memo(d['text'], d['type'], d.get('author', '研究用户'))
            m.tags = d.get('tags', [])
            return m

        for k, memos_data in data.get('doc_memos', {}).items():
            mm.doc_memos[k] = [dict_to_memo(d) for d in memos_data]
        for k, memos_data in data.get('code_memos', {}).items():
            mm.code_memos[k] = [dict_to_memo(d) for d in memos_data]
        for k, memos_data in data.get('segment_memos', {}).items():
            mm.segment_memos[k] = [dict_to_memo(d) for d in memos_data]
        mm.project_memos = [dict_to_memo(d) for d in data.get('project_memos', [])]

        return mm

    # ---- 摘要 ----
    def summary(self):
        return {
            '文档备忘录': len(self.doc_memos),
            '代码备忘录': len(self.code_memos),
            '段落备忘录': len(self.segment_memos),
            '项目备忘录': len(self.project_memos),
            '总备忘录数': sum(len(m) for m in self.doc_memos.values()) +
                        sum(len(m) for m in self.code_memos.values()) +
                        sum(len(m) for m in self.segment_memos.values()) +
                        len(self.project_memos),
            '标签': self.all_tags(),
        }
