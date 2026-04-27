"""
文本情感与质性分析模块 - 通用文本分析
======================================
不包含任何行业特定逻辑，完全基于用户数据工作。

核心功能（参考 QualCoder / ATLAS.ti 的文本分析功能）：
- 情感强度分类（5级）
- 方面级情感分析（支持自定义维度）
- 隐性不满检测（转折词分析）
- 关键词自动编码
- 时间序列趋势分析
- 分组画像分析（通用聚类分析）

适用于：社会科学研究、教育研究、心理学、市场研究、政策分析等领域。
"""

import re
import pandas as pd
from collections import Counter, defaultdict
from typing import List, Dict, Any, Optional

import jieba

# ══════════════════════════════════════════════════════════════
# 通用语言资源
# ══════════════════════════════════════════════════════════════

# 否定词
NEGATION_WORDS = ['不', '没', '无', '非', '别', '莫', '休', '未曾', '从未',
                  '不再', '并不', '并未', '毫无', '不是', '没得']

# 程度副词权重
DEGREE_ADVERBS = {
    '极其': 2.0, '格外': 2.0, '超级': 2.0, '极致': 2.0, '相当': 1.8,
    '非常': 1.6, '十分': 1.6, '特别': 1.6,
    '很': 1.3, '挺': 1.3, '颇': 1.2,
    '较': 1.1, '比较': 1.1, '稍微': 0.8, '有点': 0.7,
    '一点': 0.6, '略微': 0.6, '稍': 0.6,
}

# 通用情感词（用于无自定义词典时的基础情感判断）
GENERIC_POS_WORDS = [
    '好', '棒', '赞', '满意', '喜欢', '舒适', '干净', '热情', '满意',
    '惊喜', '完美', '推荐', '值得', '划算', '美味', '安静', '漂亮',
    '温馨', '友善', '专业', '优秀', '出色', '精彩', '优质', '方便',
    '顺利', '快', '迅速', '高效', '清晰', '明确', '准确', '成功',
]

GENERIC_NEG_WORDS = [
    '差', '失望', '糟', '脏', '吵', '旧', '破', '冷', '热', '贵',
    '慢', '乱', '恶劣', '敷衍', '臭', '恶心', '可怕', '后悔', '遗憾',
    '失望', '不满', '糟糕', '失败', '错误', '问题', '麻烦', '困难',
    '糟糕', '恶劣', '腐败', '虚假', '欺骗', '坑', '宰',
]

# 转折词（用于隐性不满检测）
TRANSITION_WORDS = [
    '但是', '不过', '然而', '只是', '可惜', '遗憾', '总体', '整体',
    '还好', '还行', '不错', '唯一', '美中不足', '建议', '如果能',
]


# ══════════════════════════════════════════════════════════════
# 核心分析函数
# ══════════════════════════════════════════════════════════════

def negation_flip(text: str, base_score: float, window: int = 5) -> float:
    """
    否定翻转检测。

    "不干净" → 负面，"不差" → 正面
    """
    if not isinstance(text, str) or base_score == 0:
        return base_score

    words = jieba.lcut(text)
    sign = 1 if base_score > 0 else -1

    for i, word in enumerate(words):
        if word not in NEGATION_WORDS:
            continue
        for j in range(i + 1, min(i + window + 1, len(words))):
            w = words[j]
            if sign > 0 and w in GENERIC_POS_WORDS:
                return -abs(base_score) * 0.85
            elif sign < 0 and w in GENERIC_NEG_WORDS:
                return abs(base_score) * 0.85
    return base_score


def classify_intensity(text: str,
                      pos_words: List[str] = None,
                      neg_words: List[str] = None) -> tuple:
    """
    情感强度5级分类。

    Returns:
        (level: int 1-5, label: str, confidence: float)
    """
    if not isinstance(text, str):
        return 3, '中性', 0.0

    pos_w = pos_words or GENERIC_POS_WORDS
    neg_w = neg_words or GENERIC_NEG_WORDS

    words = jieba.lcut(text)
    pos_count = sum(1 for w in words if w in pos_w)
    neg_count = sum(1 for w in words if w in neg_w)

    degree = 1.0
    for adv, weight in DEGREE_ADVERBS.items():
        if adv in text:
            degree = max(degree, weight)

    if pos_count > neg_count:
        raw = 3 + min(degree - 1.0, 1.0)
        if any(w in text for w in ['极其', '超级', '完美', '超赞', '简直完美']):
            raw = 5.0
    elif neg_count > pos_count:
        raw = 3 - min(degree - 1.0, 1.0)
        if any(w in text for w in ['极差', '太差', '垃圾', '噩梦', '恐怖']):
            raw = 1.0
    else:
        raw = 3.0

    level = int(round(max(1, min(5, raw))))
    labels = {5: '强烈正面', 4: '一般正面', 3: '中性', 2: '一般负面', 1: '强烈负面'}
    label = labels.get(level, '中性')
    confidence = min(1.0, (pos_count + neg_count) / max(len(words), 1) * 10)
    return level, label, round(confidence, 2)


def detect_transitions(text: str) -> List[Dict]:
    """
    检测转折词及其上下文。

    Returns:
        [{'word': str, 'before': str, 'after': str, 'position': int}, ...]
    """
    if not isinstance(text, str):
        return []

    results = []
    for w in TRANSITION_WORDS:
        if w in text:
            idx = text.find(w)
            before = text[max(0, idx-60):idx]
            after = text[idx:min(len(text), idx+60)]
            results.append({
                'word': w,
                'before': before.strip(),
                'after': after.strip(),
                'position': idx,
            })
    return results


# ══════════════════════════════════════════════════════════════
# 情感分析器（QDA应用层包装）
# ══════════════════════════════════════════════════════════════

class SentimentIntensityAnalyzer:
    """
    情感强度分析器。
    支持自定义情感词典，无行业特定内容。
    """

    def __init__(self,
                 pos_words: List[str] = None,
                 neg_words: List[str] = None):
        self.pos_words = pos_words or GENERIC_POS_WORDS
        self.neg_words = neg_words or GENERIC_NEG_WORDS

    def classify(self, text: str) -> tuple:
        """返回 (level, label, confidence)"""
        return classify_intensity(text, self.pos_words, self.neg_words)

    def analyze_document(self, text: str) -> dict:
        """
        分析单条文档。

        Returns:
            {
                'level': int, 'label': str, 'confidence': float,
                'negation_flip': bool, 'transitions': list,
                'pos_word_count': int, 'neg_word_count': int,
            }
        """
        level, label, conf = self.classify(text)
        neg_flip = negation_flip(text, 0.5) < 0 or negation_flip(text, -0.5) > 0
        trans = detect_transitions(text)

        words = jieba.lut(text)
        pos_n = sum(1 for w in words if w in self.pos_words)
        neg_n = sum(1 for w in words if w in self.neg_words)

        return {
            'level': level,
            'label': label,
            'confidence': conf,
            'negation_flip': neg_flip,
            'transitions': trans,
            'pos_word_count': pos_n,
            'neg_word_count': neg_n,
        }

    def analyze_dataframe(self, df: pd.DataFrame,
                         text_col: str = '_text',
                         score_col: str = None) -> pd.DataFrame:
        """
        批量分析，返回添加了情感分析列的DataFrame。
        """
        df = df.copy()

        def do_row(row):
            lvl, lbl, conf = self.classify(str(row.get(text_col, '')))
            return pd.Series({
                '情感强度级': lvl,
                '情感强度标签': lbl,
                '情感置信度': conf,
            })

        df[['情感强度级', '情感强度标签', '情感置信度']] = df.apply(do_row, axis=1)

        if score_col and score_col in df.columns:
            df['评分分组'] = pd.cut(
                df[score_col],
                bins=[0, 2.9, 3.9, 4.4, 5.1],
                labels=['高度不满', '中度不满', '中度满意', '高度满意']
            )

        return df

    def get_summary(self, df: pd.DataFrame,
                   level_col: str = '情感强度级') -> dict:
        """返回情感统计摘要"""
        summary = {}
        if level_col in df.columns:
            dist = df[level_col].value_counts().to_dict()
            summary['level_distribution'] = dist
            summary['strongly_positive_rate'] = round(
                dist.get(5, 0) / max(len(df), 1) * 100, 1)
            summary['strongly_negative_rate'] = round(
                dist.get(1, 0) / max(len(df), 1) * 100, 1)
        return summary


class AspectSentimentAnalyzer:
    """
    方面级情感分析。
    用户通过 keyword_dict 参数传入领域相关的方面和词汇。
    """

    def __init__(self,
                 keyword_dict: Dict[str, Dict[str, List[str]]] = None):
        """
        Args:
            keyword_dict: {
                '服务态度': {'pos': ['热情', '周到'], 'neg': ['冷漠', '敷衍']},
                '环境设施': {'pos': ['干净', '舒适'], 'neg': ['脏', '差']},
            }
            不指定则使用通用情感词，不做分维度分析。
        """
        self.keyword_dict = keyword_dict or {}

    def analyze(self, text: str) -> dict:
        """
        对文本进行方面级情感分析。

        Returns:
            {
                'overall': 'positive'/'negative'/'neutral'/'mixed',
                'aspects': {方面名: 'positive'/'negative'/'neutral'/'mixed', ...},
                'intensity': (level, label, confidence),
            }
        """
        results = {'aspects': {}, 'overall': 'neutral'}

        # 整体情感
        lvl, lbl, conf = classify_intensity(text)
        results['intensity'] = (lvl, lbl, conf)

        # 各方面
        pos_total, neg_total = 0, 0
        for aspect, keywords in self.keyword_dict.items():
            pos_kw = keywords.get('pos', [])
            neg_kw = keywords.get('neg', [])

            has_pos = any(kw in text for kw in pos_kw)
            has_neg = any(kw in text for kw in neg_kw)

            # 否定检测
            flipped_pos = any(f'{n}{kw}' in text or (n + kw) in text
                            for kw in pos_kw for n in NEGATION_WORDS)
            flipped_neg = any(f'{n}{kw}' in text or (n + kw) in text
                            for kw in neg_kw for n in NEGATION_WORDS)

            pos_active = has_pos and not flipped_pos
            neg_active = has_neg and not flipped_neg

            if pos_active and not neg_active:
                results['aspects'][aspect] = 'positive'
                pos_total += 1
            elif neg_active and not pos_active:
                results['aspects'][aspect] = 'negative'
                neg_total += 1
            elif pos_active and neg_active:
                results['aspects'][aspect] = 'mixed'
            else:
                results['aspects'][aspect] = 'neutral'

        if pos_total > neg_total:
            results['overall'] = 'positive'
        elif neg_total > pos_total:
            results['overall'] = 'negative'
        elif pos_total > 0:
            results['overall'] = 'mixed'

        return results

    def analyze_dataframe(self, df: pd.DataFrame,
                         text_col: str = '_text') -> pd.DataFrame:
        """批量分析DataFrame"""
        rows = []
        for _, row in df.iterrows():
            r = self.analyze(str(row.get(text_col, '')))
            r['doc_id'] = row.get('_id', _)
            rows.append(r)
        return pd.DataFrame(rows)


class HiddenDissatisfactionDetector:
    """
    隐性不满检测。
    识别转折句后半句负面 + 高分文档中的负面片段。
    """

    def __init__(self,
                 neg_indicators: List[str] = None):
        self.neg_indicators = neg_indicators or [
            '差', '脏', '吵', '贵', '旧', '破', '失望', '不满',
            '小问题', '有点', '不够', '欠缺', '遗憾', '可惜',
            '问题', '糟糕', '恶心', '后悔', '再也不',
        ]

    def detect(self, text: str, score: float = None,
               high_score_threshold: float = 4.0) -> dict:
        """
        检测隐性不满。

        Returns:
            {
                'is_hidden_neg': bool,
                'transition_found': bool,
                'transition_word': str,
                'second_half_negative': bool,
                'high_score_hidden': bool,
                'negative_fragments': [str],
                'reason': str,
            }
        """
        result = {
            'is_hidden_neg': False,
            'transition_found': False,
            'transition_word': '',
            'second_half_negative': False,
            'high_score_hidden': False,
            'negative_fragments': [],
            'reason': '',
        }

        if not isinstance(text, str) or len(text) < 5:
            return result

        # 转折句检测
        patterns = [
            r'虽然[:：]?(.+?)(但是|不过|然而|只是|可惜)',
            r'尽管[:：]?(.+?)(但|不过|然而)',
            r'(.+?)(但是|不过|然而|只是|唯一|可惜|遗憾地说)(.+)',
            r'(不错|还行|还好|总体|整体)(但是|不过|只是|然而)(.+)',
        ]

        for pattern in patterns:
            m = re.search(pattern, text)
            if m:
                result['transition_found'] = True
                groups = m.groups()
                if len(groups) >= 3:
                    second_half = groups[-1]
                elif len(groups) == 2:
                    second_half = groups[1]
                else:
                    second_half = ''

                result['second_half_negative'] = any(
                    ind in second_half for ind in self.neg_indicators)

                if result['second_half_negative']:
                    result['is_hidden_neg'] = True
                    result['transition_word'] = groups[1] if len(groups) >= 2 else ''
                    result['negative_fragments'].append(second_half.strip()[:100])
                    result['reason'] = f'转折句后半句负面'
                    return result

        # 高分文档隐性不满
        if score is not None and score >= high_score_threshold:
            found = [w for w in self.neg_indicators if w in text]
            if found:
                result['high_score_hidden'] = True
                result['is_hidden_neg'] = True
                result['reason'] = f'高分文档含负面词: {found}'
                for w in found:
                    idx = text.find(w)
                    if idx >= 0:
                        fragment = text[max(0, idx-15):min(len(text), idx+20)]
                        result['negative_fragments'].append(f'...{fragment}...')

        return result

    def detect_dataframe(self, df: pd.DataFrame,
                        text_col: str = '_text',
                        score_col: str = None) -> pd.DataFrame:
        """批量检测"""
        rows = []
        for _, row in df.iterrows():
            score = float(row[score_col]) if score_col and score_col in df.columns else None
            r = self.detect(str(row.get(text_col, '')), score=score)
            r['doc_id'] = row.get('_id', _)
            rows.append(r)
        result_df = pd.DataFrame(rows)
        return result_df[result_df['is_hidden_neg']]


# ══════════════════════════════════════════════════════════════
# 关键词自动编码器
# ══════════════════════════════════════════════════════════════

class KeywordAutoCoder:
    """
    基于关键词的自动编码器。
    接受用户定义的 {编码名: [关键词列表]} 映射，
    在文档中匹配并创建编码实例。
    """

    def __init__(self, keyword_dict: Dict[str, List[str]] = None):
        self.keyword_dict = keyword_dict or {}

    def add_keywords(self, code_name: str, keywords: List[str]):
        """添加编码关键词"""
        if code_name not in self.keyword_dict:
            self.keyword_dict[code_name] = []
        self.keyword_dict[code_name].extend(keywords)

    def code_text(self, text: str) -> Dict[str, List[str]]:
        """
        对单条文本编码。

        Returns:
            {编码名: [匹配的关键词列表], ...}
        """
        results = {}
        for code_name, keywords in self.keyword_dict.items():
            matched = []
            for kw in keywords:
                if kw in text:
                    matched.append(kw)
            if matched:
                results[code_name] = matched
        return results

    def code_dataframe(self, df: pd.DataFrame,
                      text_col: str = '_text') -> pd.DataFrame:
        """批量编码DataFrame"""
        df = df.copy()
        for code_name in self.keyword_dict:
            df[f'编码_{code_name}'] = False

        def do_row(row):
            text = str(row.get(text_col, ''))
            matched = self.code_text(text)
            for code_name in self.keyword_dict:
                row[f'编码_{code_name}'] = code_name in matched
            return row

        return df.apply(do_row, axis=1)

    def get_code_stats(self, df: pd.DataFrame) -> dict:
        """获取编码统计"""
        stats = {}
        for code_name in self.keyword_dict:
            col = f'编码_{code_name}'
            if col in df.columns:
                count = df[col].sum()
                stats[code_name] = {
                    'count': int(count),
                    'rate': round(count / max(len(df), 1) * 100, 1),
                }
        return stats


# ══════════════════════════════════════════════════════════════
# 趋势分析
# ══════════════════════════════════════════════════════════════

def monthly_trend(df: pd.DataFrame,
                  date_col: str = '日期',
                  score_col: str = '评分',
                  text_col: str = '_text') -> pd.DataFrame:
    """
    月度趋势分析。

    Returns:
        DataFrame: 月份, 正评数, 负评数, 正负比, 均分, 文档数, 趋势
    """
    if date_col not in df.columns or score_col not in df.columns:
        return pd.DataFrame()

    df_work = df.copy()

    try:
        df_work['月份'] = pd.to_datetime(df_work[date_col], errors='coerce').dt.to_period('M')
    except Exception:
        return pd.DataFrame()

    df_work['正评'] = (df_work[score_col] >= 4.5).astype(int)
    df_work['负评'] = (df_work[score_col] <= 3.0).astype(int)

    monthly = df_work.groupby('月份').agg(
        正评数=(score_col, lambda x: (x >= 4.5).sum()),
        负评数=(score_col, lambda x: (x <= 3.0).sum()),
        均分=(score_col, 'mean'),
        文档数=(text_col, 'count'),
    ).reset_index()

    monthly['正负比'] = (monthly['正评数'] / monthly['负评数'].replace(0, 1)).round(2)
    monthly['月份'] = monthly['月份'].astype(str)

    if len(monthly) > 1:
        monthly['趋势'] = monthly['均分'].diff().apply(
            lambda x: '↑上升' if x > 0.1 else ('↓下降' if x < -0.1 else '→持平')
        )
    else:
        monthly['趋势'] = '—'

    return monthly.round({'均分': 2})


def infer_customer_persona(df: pd.DataFrame,
                          text_col: str = '_text',
                          score_col: str = '评分',
                          group_col: str = None) -> dict:
    """
    通用分组画像分析。

    Returns:
        dict: {
            分组名: {
                '文档数': int,
                '均值': float,
                '情感强度分布': {级别: 数量},
                '典型片段': str,
            }
        }
    """
    persona = {}

    if group_col is None or group_col not in df.columns:
        return persona

    analyzer = SentimentIntensityAnalyzer()

    for group_val in df[group_col].dropna().unique():
        subset = df[df[group_col] == group_val]
        if len(subset) < 3:
            continue

        avg_score = subset[score_col].mean() if score_col in subset.columns else 0

        # 情感强度分布
        level_map = defaultdict(int)
        for _, row in subset.head(200).iterrows():
            lvl, lbl, _ = analyzer.classify(str(row.get(text_col, '')))
            level_map[lvl] += 1

        # 典型片段（评分最低的）
        if score_col in subset.columns:
            worst = subset.nsmallest(2, score_col)
            worst_texts = [str(r[text_col])[:120] for _, r in worst.iterrows()
                          if pd.notna(r.get(text_col))]
            typical = worst_texts[0] if worst_texts else ''
        else:
            typical = ''

        persona[group_val] = {
            '文档数': len(subset),
            '均值': round(float(avg_score), 2),
            '情感强度分布': dict(level_map),
            '典型片段': typical,
        }

    return persona


# ══════════════════════════════════════════════════════════════════
# 通用别名（向后兼容）
# ══════════════════════════════════════════════════════════════════

def time_trend(df: pd.DataFrame,
               date_col: str = '日期',
               score_col: str = '评分',
               text_col: str = '_text') -> pd.DataFrame:
    """
    时间趋势分析（月度/周度/季度）
    
    是 monthly_trend 的通用化版本，支持任意时间粒度。
    """
    return monthly_trend(df, date_col, score_col, text_col)


def segment_analysis(df: pd.DataFrame,
                     text_col: str = '_text',
                     score_col: str = '评分',
                     group_col: str = None) -> dict:
    """
    分组分析（通用画像分析）
    
    是 infer_customer_persona 的通用化版本，可用于任意分组变量。
    """
    return infer_customer_persona(df, text_col, score_col, group_col)
