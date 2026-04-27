"""
可视化模块 - 通用图表生成器
支持：matplotlib图表（存储为图片）、HTML交互图表
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from pathlib import Path
import base64
from io import BytesIO

# 中文字体
try:
    plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'Arial Unicode MS']
except:
    pass
plt.rcParams['axes.unicode_minus'] = False


# ==================== 颜色方案 ====================

POS_COLOR = '#43a047'
NEG_COLOR = '#e53935'

# 情感金字塔默认关键词（可由调用者自定义）
SENTIMENT_POS_KW = ['非常', '特别', '十分', '超级', '完美', '惊喜', '极其', '优秀']
SENTIMENT_NEG_KW = ['太差', '极差', '垃圾', '恶心', '恐怖', '恶劣', '失望', '糟糕']
NEU_COLOR = '#888888'
GROUP_COLORS = [
    '#1976d2', '#d81b60', '#7b1fa2', '#388e3c',
    '#f57c00', '#0097a7', '#5d4037', '#455a64',
]


# ==================== 图表基类 ====================

class Chart:
    """单个图表基类"""

    def __init__(self, figsize=(10, 6), dpi=120):
        self.figsize = figsize
        self.dpi = dpi
        self.fig = None
        self.ax = None

    def save(self, path):
        if self.fig:
            self.fig.savefig(path, dpi=self.dpi, bbox_inches='tight',
                          facecolor='white', edgecolor='none')
            return path

    def to_base64(self):
        """转为base64用于HTML嵌入"""
        if not self.fig:
            return ''
        buf = BytesIO()
        self.fig.savefig(buf, format='png', dpi=self.dpi,
                        bbox_inches='tight', facecolor='white')
        buf.seek(0)
        return base64.b64encode(buf.read()).decode()


# ==================== 1. 雷达图：多维度对比 ====================

def plot_dimension_radar(df_list, group_names, dims, title='维度对比雷达图'):
    """
    雷达图：展示多个组的各维度正负提及率

    Args:
        df_list: [df1, df2, ...]
        group_names: ['组A', '组B', ...]
        dims: 维度列表

    Returns:
        Chart对象
    """
    N = len(dims)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]  # 闭合

    fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True), dpi=120)

    for df, name, color in zip(df_list, group_names, GROUP_COLORS[:len(df_list)]):
        values_pos = []
        values_neg = []
        for dim in dims:
            pos_col = f'代码_{dim}_正'
            neg_col = f'代码_{dim}_负'
            if pos_col in df.columns:
                total = df[pos_col].sum() + df[neg_col].sum()
                pos = df[pos_col].sum()
                neg = df[neg_col].sum()
            else:
                total, pos, neg = 1, 0, 0
            values_pos.append(pos / max(total, 1) * 100)
            values_neg.append(neg / max(total, 1) * 100)

        values_pos += values_pos[:1]
        values_neg += values_neg[:1]

        ax.plot(angles, values_pos, 'o-', linewidth=2, label=f'{name} 正面',
                color=color, alpha=0.8)
        ax.fill(angles, values_pos, alpha=0.1, color=color)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(dims, size=9)
    ax.set_title(title, size=13, pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
    ax.set_ylim(0, 50)

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 2. 堆叠柱状图：维度正负对比 ====================

def plot_dimension_bar(df_list, group_names, dims, title='维度正负对比'):
    """
    分组柱状图：各维度正面vs负面编码数
    """
    n_dims = len(dims)
    n_groups = len(df_list)
    x = np.arange(n_dims)
    width = 0.6 / n_groups

    fig, ax = plt.subplots(figsize=(max(12, n_dims * 1.5), 7), dpi=120)

    for i, (df, name) in enumerate(zip(df_list, group_names)):
        pos_vals = []
        neg_vals = []
        for dim in dims:
            pos_col = f'代码_{dim}_正'
            neg_col = f'代码_{dim}_负'
            pos_vals.append(df[pos_col].sum() if pos_col in df.columns else 0)
            neg_vals.append(-df[neg_col].sum() if neg_col in df.columns else 0)

        offset = (i - n_groups / 2 + 0.5) * width
        bars_pos = ax.bar(x + offset, pos_vals, width, label=f'{name} 正面',
                          color=GROUP_COLORS[i], alpha=0.8)
        ax.bar(x + offset, neg_vals, width, label=f'{name} 负面',
               color=GROUP_COLORS[i], alpha=0.4, hatch='//')

    ax.set_xlabel('维度')
    ax.set_ylabel('编码数量（正面↑ 负面↓）')
    ax.set_title(title, size=13)
    ax.set_xticks(x)
    ax.set_xticklabels(dims, rotation=30, ha='right', size=8)
    ax.legend(fontsize=8)
    ax.axhline(0, color='black', linewidth=0.5)
    ax.grid(axis='y', alpha=0.3)

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 3. 热力图：共现矩阵 ====================

def plot_cooccurrence_heatmap(matrix, title='维度共现热力图', figsize=(10, 8)):
    """
    热力图：展示维度共现关系
    """
    try:
        import seaborn as sns
        has_seaborn = True
    except ImportError:
        has_seaborn = False

    fig, ax = plt.subplots(figsize=figsize, dpi=120)

    if has_seaborn and not matrix.empty:
        sns.heatmap(matrix, annot=True, fmt='d', cmap='Blues',
                    ax=ax, cbar_kws={'label': '共现次数'},
                    linewidths=0.5)
    else:
        ax.imshow(matrix.values, cmap='Blues', aspect='auto')
        ax.set_xticks(range(len(matrix.columns)))
        ax.set_yticks(range(len(matrix.index)))
        ax.set_xticklabels(matrix.columns, rotation=45, ha='right', size=8)
        ax.set_yticklabels(matrix.index, size=8)

    ax.set_title(title, size=13)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 4. 属性分布对比图 ====================

def plot_attribute_distribution(df_list, group_names, column_name, title='属性分布对比'):
    """
    分组柱状图：某属性在各取值上的分布对比

    Args:
        df_list: [df1, df2, ...]
        group_names: ['组A', '组B', ...]
        column_name: 要分析的列名
        title: 图表标题
    """
    fig, ax = plt.subplots(figsize=(8, 6), dpi=120)

    # 收集所有唯一取值
    all_vals = set()
    for df in df_list:
        if column_name in df.columns:
            all_vals.update(df[column_name].dropna().unique())
    all_vals = sorted(all_vals)

    n_vals = len(all_vals)
    if n_vals == 0:
        ax.text(0.5, 0.5, f'列 "{column_name}" 无有效数据', ha='center', va='center', transform=ax.transAxes)
        ax.set_title(title, size=13)
        plt.tight_layout()
        chart = Chart()
        chart.fig = fig
        chart.ax = ax
        return chart

    x = np.arange(n_vals)
    n_groups = len(df_list)
    width = 0.6 / max(n_groups, 1)

    for i, (df, name) in enumerate(zip(df_list, group_names)):
        if column_name in df.columns:
            dist = df[column_name].value_counts()
            vals = [dist.get(s, 0) for s in all_vals]
        else:
            vals = [0] * n_vals
        total = sum(vals)
        pcts = [v / max(total, 1) * 100 for v in vals]

        offset = (i - n_groups / 2 + 0.5) * width
        ax.bar(x + offset, pcts, width, label=name, color=GROUP_COLORS[i], alpha=0.8)

    ax.set_xlabel(column_name)
    ax.set_ylabel('占比（%）')
    ax.set_title(title, size=13)
    ax.set_xticks(x)
    ax.set_xticklabels([str(v) for v in all_vals])
    ax.legend()
    ax.grid(axis='y', alpha=0.3)

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 5. 情感金字塔：情感强度分层 ====================

def plot_sentiment_pyramid(df, text_col='文本', score_col='分值', title='情感强度金字塔',
                           pos_keywords=None, neg_keywords=None):
    """
    金字塔图：强正面 → 中性 → 强负面

    Args:
        df: 数据框
        text_col: 文本内容列名
        score_col: 分数/评分列名
        title: 图表标题
        pos_keywords: 自定义正面关键词列表（默认使用 SENTIMENT_POS_KW）
        neg_keywords: 自定义负面关键词列表（默认使用 SENTIMENT_NEG_KW）

    Returns:
        Chart对象
    """
    fig, ax = plt.subplots(figsize=(9, 7), dpi=120)

    pos_kw = pos_keywords if pos_keywords is not None else SENTIMENT_POS_KW
    neg_kw = neg_keywords if neg_keywords is not None else SENTIMENT_NEG_KW

    # 情感分类
    def classify(text):
        if not isinstance(text, str):
            return 'neutral'
        pos = any(k in text for k in pos_kw)
        neg = any(k in text for k in neg_kw)
        if pos and neg:
            return 'mixed'
        elif pos:
            return 'strong_pos'
        elif neg:
            return 'strong_neg'
        return 'mild'

    labels = ['强正面', '中度正面', '中性', '中度负面', '强负面']
    colors = ['#1b5e20', '#66bb6a', '#9e9e9e', '#ef5350', '#b71c1c']
    sizes = [0, 0, 0, 0, 0]

    for _, row in df.iterrows():
        cls = classify(row.get(text_col, ''))
        score = row.get(score_col, 3)
        if cls == 'strong_pos':
            sizes[0] += 1
        elif cls == 'mild' and score >= 4:
            sizes[1] += 1
        elif cls == 'mild' and score == 3:
            sizes[2] += 1
        elif cls == 'mild' and score < 3:
            sizes[3] += 1
        elif cls == 'strong_neg' or score < 2.5:
            sizes[4] += 1

    y_pos = range(len(labels))
    max_size = max(sizes) if max(sizes) > 0 else 1

    for i, (label, size, color) in enumerate(zip(labels, sizes, colors)):
        pct = size / len(df) * 100
        width = size / max_size * 10
        ax.barh(y_pos[i], width, color=color, alpha=0.8, height=0.6)
        ax.text(width + 0.1, y_pos[i], f'{size}条 ({pct:.1f}%)',
                va='center', fontsize=10)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels)
    ax.set_xlabel('情感强度')
    ax.set_title(title, size=13)
    ax.set_xlim(0, 12)
    ax.invert_yaxis()

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 词云图 ====================

def plot_wordcloud(word_freq: dict,
                   title: str = '词云图',
                   width: int = 1200,
                   height: int = 600,
                   max_words: int = 200,
                   background_color: str = 'white',
                   colormap: str = 'viridis',
                   font_path: str = None) -> Chart:
    """
    生成词云图。

    Parameters
    ----------
    word_freq : dict  {词项: 词频或TF-IDF值, ...}
    title : str  图表标题
    width, height : int  画布尺寸（像素）
    max_words : int  最大词数
    background_color : str  背景色
    colormap : str  matplotlib 色彩映射
    font_path : str  中文.ttf 字体路径（可选，自动检测）

    Returns
    -------
    Chart对象
    """
    try:
        from wordcloud import WordCloud
    except ImportError:
        raise ImportError("请安装 wordcloud：pip install wordcloud")

    import os as _os
    if font_path is None:
        font_candidates = [
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
            '/System/Library/Fonts/PingFang.ttc',
            'C:/Windows/Fonts/msyh.ttc',
        ]
        for fp in font_candidates:
            if _os.path.exists(fp):
                font_path = fp
                break

    wc = WordCloud(
        font_path=font_path,
        width=width,
        height=height,
        max_words=max_words,
        background_color=background_color,
        colormap=colormap,
        prefer_horizontal=0.7,
        min_font_size=10,
        max_font_size=150,
        scale=1.5,
    )
    wc.generate_from_frequencies(word_freq)

    fig, ax = plt.subplots(figsize=(width / 80, height / 80), dpi=80)
    ax.imshow(wc, interpolation='bilinear')
    ax.axis('off')
    ax.set_title(title, size=14, pad=10)

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 词频柱状图 ====================

def plot_word_frequency_barchart(word_freq_df,
                                 top_n: int = 30,
                                 title: str = '词频统计（Top N）',
                                 color: str = '#1565c0') -> Chart:
    """
    词频水平柱状图（支持原始词频或TF-IDF）。
    """
    if 'count' in word_freq_df.columns:
        sorted_col = 'count'
    else:
        num_cols = word_freq_df.select_dtypes('number').columns
        sorted_col = num_cols[0] if len(num_cols) > 0 else word_freq_df.columns[0]

    top = word_freq_df[sorted_col].nlargest(top_n)
    labels = list(top.index)
    values = list(top.values)

    fig, ax = plt.subplots(figsize=(10, max(6, top_n * 0.35)), dpi=120)
    bars = ax.barh(range(len(labels)), values, color=color, alpha=0.85)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel('词频' if sorted_col == 'count' else 'TF-IDF 值')
    ax.set_title(title, size=13)
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values) * 0.005,
                bar.get_y() + bar.get_height() / 2,
                f'{int(val):,}', va='center', fontsize=8)
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== N-gram 分布图 ====================

def plot_ngram_distribution(ngram_counts: dict,
                             title: str = 'N-gram 词组分布') -> Chart:
    """
    N-gram 词组频率水平条形图。
    """
    sorted_items = sorted(ngram_counts.items(), key=lambda x: -x[1])[:30]
    labels = [k for k, _ in sorted_items]
    values = [v for _, v in sorted_items]

    fig, ax = plt.subplots(figsize=(10, max(5, len(labels) * 0.4)), dpi=120)
    bars = ax.barh(range(len(labels)), values, color='#7b1fa2', alpha=0.8)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel('出现次数')
    ax.set_title(title, size=13)
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values or [1]) * 0.005,
                bar.get_y() + bar.get_height() / 2,
                f'{int(val):,}', va='center', fontsize=8)
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 文档覆盖率图 ====================

def plot_document_coverage(word_doc_df,
                           title: str = '词项文档覆盖率（Top 30）') -> Chart:
    """
    每个词出现在多少个文档中的条形图。
    """
    doc_coverage = (word_doc_df > 0).sum(axis=1).nlargest(30)
    labels = list(doc_coverage.index)
    values = list(doc_coverage.values)
    n_docs = word_doc_df.shape[1]

    fig, ax = plt.subplots(figsize=(10, max(5, len(labels) * 0.4)), dpi=120)
    bars = ax.barh(range(len(labels)), values, color='#00695c', alpha=0.8)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel(f'出现文档数（共 {n_docs} 篇）')
    ax.set_title(title, size=13)
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + 0.3,
                bar.get_y() + bar.get_height() / 2,
                f'{int(val)}', va='center', fontsize=8)
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 6. 矩阵热力图：交叉分析 ====================

def plot_cross_heatmap(matrix_dict, title='交叉分析热力图'):
    """
    分组与维度的交叉热力图组
    """
    n = len(matrix_dict)
    if n == 0:
        return None

    fig, axes = plt.subplots(1, n, figsize=(7 * n, 7), dpi=100)
    if n == 1:
        axes = [axes]

    try:
        import seaborn as sns
        has_sns = True
    except ImportError:
        has_sns = False

    for ax, (dim, mat) in zip(axes, matrix_dict.items()):
        if has_sns and not mat.empty:
            sns.heatmap(mat, annot=True, fmt='.0f', cmap='Greens',
                       ax=ax, cbar_kws={'shrink': 0.8})
        ax.set_title(dim, size=11)
        ax.set_xlabel('分组')

    fig.suptitle(title, size=13, y=1.02)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = axes
    return chart


# ==================== 7. 代码分布饼图 ====================

def plot_code_distribution(df, dims, code_prefix='代码_', title='维度分布', top_n=8):
    """
    饼图：各维度编码占比
    """
    sizes = []
    labels = []
    for dim in dims[:top_n]:
        pos_col = f'{code_prefix}{dim}_正'
        neg_col = f'{code_prefix}{dim}_负'
        total = df[pos_col].sum() + df[neg_col].sum() if pos_col in df.columns else 0
        if total > 0:
            sizes.append(total)
            labels.append(dim)

    if not sizes:
        return None

    fig, ax = plt.subplots(figsize=(9, 9), dpi=120)

    colors = GROUP_COLORS[:len(sizes)]
    explode = [0.03] * len(sizes)

    wedges, texts, autotexts = ax.pie(
        sizes, labels=labels, colors=colors, explode=explode,
        autopct='%1.1f%%', startangle=90,
        pctdistance=0.75, labeldistance=1.1
    )
    for t in texts:
        t.set_fontsize(9)
    for at in autotexts:
        at.set_fontsize(8)

    ax.set_title(title, size=13)
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 8. 时间序列趋势图 ====================

def plot_monthly_trend(dfs_dict, metric='分值', title='月度趋势',
                       date_col='日期', score_col='分值', text_col='文本'):
    """
    折线图：多个组的月度指标趋势

    Args:
        dfs_dict: {group_name: df}
        metric: 要展示的指标名
        date_col: 日期列名
        score_col: 分数列名
        text_col: 文本列名
    """
    import pandas as pd

    fig, ax = plt.subplots(figsize=(12, 6), dpi=120)

    color_iter = iter(GROUP_COLORS)
    for name, df in dfs_dict.items():
        color = next(color_iter)
        dfp = df.copy()
        if '月份' not in dfp.columns:
            if date_col in dfp.columns:
                dfp['月份'] = pd.to_datetime(dfp[date_col], errors='coerce').dt.to_period('M')
            else:
                continue

        if '月份' not in dfp.columns:
            continue

        monthly_vals = dfp.groupby('月份').size()
        ax.plot(monthly_vals.index.astype(str), monthly_vals.values, 'o-',
               color=color, linewidth=2, label=name)
        ax.set_ylabel('文档数量')

    ax.set_xlabel('月份')
    ax.set_title(title, size=13)
    ax.legend()
    ax.grid(alpha=0.3)

    # x轴标签旋转
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()

    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 9. 分组对比雷达图 ====================

def plot_group_comparison_radar(df_primary, df_secondary, name_a, name_b, dims):
    """
    两组数据对比雷达图
    """
    N = len(dims)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(9, 9), subplot_kw=dict(polar=True), dpi=120)

    for df, name, color in [(df_primary, name_a, '#1976d2'),
                              (df_secondary, name_b, '#e53935')]:
        pos_vals = []
        for dim in dims:
            pos_col = f'代码_{dim}_正'
            neg_col = f'代码_{dim}_负'
            total = df[pos_col].sum() + df[neg_col].sum() if pos_col in df.columns else 1
            pos = df[pos_col].sum() if pos_col in df.columns else 0
            pos_vals.append(pos / max(total, 1) * 100)

        pos_vals += pos_vals[:1]
        ax.plot(angles, pos_vals, 'o-', linewidth=2, label=name, color=color)
        ax.fill(angles, pos_vals, alpha=0.1, color=color)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(dims, size=9)
    ax.set_title(f'{name_a} vs {name_b} 对比', size=13, pad=20)
    ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
    ax.set_ylim(0, 50)

    plt.tight_layout()
    chart = Chart()
    chart.fig = fig
    chart.ax = ax
    return chart


# ==================== 10. 汇总图表生成器 ====================

def generate_all_charts(df, df_secondary=None,
                        group_name='组A', compare_name='组B',
                        code_prefix='代码_'):
    """
    为单个组生成所有图表

    Returns:
        dict: {chart_name: Chart}
    """
    charts = {}

    # 获取维度列表
    dims = sorted([c.replace(code_prefix, '').replace('_正', '')
                   for c in df.columns if c.endswith('_正') and c.startswith(code_prefix)])

    # 1. 维度分布饼图
    pie = plot_code_distribution(df, dims, code_prefix=code_prefix,
                                 title=f'{group_name} 维度分布')
    if pie:
        charts['维度分布饼图'] = pie

    # 2. 属性分布
    attr_chart = plot_attribute_distribution([df], [group_name], column_name='',
                                              title=f'{group_name} 属性分布')
    charts['属性分布'] = attr_chart

    # 3. 情感金字塔
    pyr = plot_sentiment_pyramid(df, title=f'{group_name} 情感强度分布')
    charts['情感金字塔'] = pyr

    # 4. 分组对比雷达图（如果有）
    if df_secondary is not None:
        radar = plot_group_comparison_radar(df, df_secondary,
                                       group_name, compare_name, dims)
        charts['对比雷达图'] = radar

    return charts


# ==================== HTML图表导出 ====================

def charts_to_html(charts_dict, title='分析图表'):
    """
    将Chart对象组转换为可嵌入HTML的base64图片
    用于嵌入GUI或网页
    """
    html_parts = [f'<h2>{title}</h2>']
    for name, chart in charts_dict.items():
        img_data = chart.to_base64()
        if img_data:
            html_parts.append(f'<h3>{name}</h3>')
            html_parts.append(f'<img src="data:image/png;base64,{img_data}" />')
    return '\n'.join(html_parts)
