"""
酒店评论综合分析：佰翔酒店 vs 漳州宾馆
"""

import pandas as pd
import jieba
from collections import Counter, defaultdict

bx = pd.read_excel('佰翔酒店合并数据.xlsx')
zz = pd.read_excel('漳州宾馆合并数据.xlsx')

for df, name in [(bx, '佰翔'), (zz, '漳州')]:
    df['评论长度'] = df['评论内容'].apply(lambda x: len(str(x)) if isinstance(x, str) else 0)
    df['有效评论'] = df['评论长度'] > 10
    df['月份'] = pd.to_datetime(df['评论日期'], errors='coerce').dt.to_period('M')

coding_dict = {
    '清洁卫生': {
        'pos': ['干净', '整洁', '卫生', '清洁', '无灰尘', '清新'],
        'neg': ['脏', '有灰尘', '毛发', '霉味', '异味', '不干净', '恶心', '有虫']
    },
    '服务态度': {
        'pos': ['热情', '周到', '贴心', '耐心', '友好', '温暖', '亲切', '礼貌', '主动', '微笑', '细心'],
        'neg': ['冷淡', '冷漠', '敷衍', '态度差', '爱答不理', '不耐烦', '差劲', '恶劣', '切单', '看人下菜']
    },
    '前台服务': {
        'pos': ['前台热情', '办理快', '效率高', '入住快', '服务好', '经理', '主管'],
        'neg': ['前台差', '入住慢', '等很久', '办理慢', '服务差', '前台冷漠']
    },
    '设施设备': {
        'pos': ['设施齐全', '设备好', '新', '现代化', '配置好', '完善'],
        'neg': ['设施旧', '设备差', '老旧', '陈旧', '坏', '损坏', '故障', '不能用']
    },
    '房间条件': {
        'pos': ['房间大', '宽敞', '舒适', '床舒服', '床软', '床品好', '采光好', '空调好', '热水好'],
        'neg': ['房间小', '狭小', '窄', '拥挤', '床硬', '不舒服', '空调差', '冷', '热']
    },
    '隔音效果': {
        'pos': ['隔音好', '安静', '很静', '噪音小'],
        'neg': ['隔音差', '吵', '噪音大', '很吵', '嘈杂', '车声', '喇叭', '施工', '装修']
    },
    '早餐餐饮': {
        'pos': ['早餐好', '丰富', '品种多', '味道好', '好吃', '中式', '西式', '自助'],
        'neg': ['早餐差', '单一', '品种少', '难吃', '冷', '不好', '没早餐']
    },
    '位置交通': {
        'pos': ['位置好', '方便', '近', '交通便利', '市中心', '周边有', '万达', '地铁'],
        'neg': ['位置偏', '偏远', '不方便', '难找', '偏僻']
    },
    '性价比': {
        'pos': ['物超所值', '划算', '值得', '超值', '性价比高', '价格合理'],
        'neg': ['不值', '贵', '性价比低', '价格高', '太贵', '坑']
    },
    '景观环境': {
        'pos': ['景观好', '风景美', '江景', '花园', '漂亮', '美', '景色好'],
        'neg': ['景观差', '难看', '看不到', '阴森']
    },
    '停车配套': {
        'pos': ['停车方便', '有停车场', '车位多', '免费停车'],
        'neg': ['停车难', '没车位', '收费', '停车贵']
    }
}

def code_review(text):
    if not isinstance(text, str):
        return [], []
    pos_dims = []
    neg_dims = []
    for dim, words in coding_dict.items():
        for w in words['pos']:
            if w in text:
                pos_dims.append(dim)
                break
        for w in words['neg']:
            if w in text:
                neg_dims.append(dim)
                break
    return list(set(pos_dims)), list(set(neg_dims))

def analyze_hotel(df, name):
    sep = '='*60
    print(sep)
    print(name + '酒店 - 主题分析与情感分析')
    print(sep)

    pos_count = defaultdict(int)
    neg_count = defaultdict(int)
    neg_examples = defaultdict(list)

    for _, row in df[df['有效评论']].iterrows():
        text = str(row['评论内容'])
        pos_dims, neg_dims = code_review(text)
        for d in pos_dims:
            pos_count[d] += 1
        for d in neg_dims:
            neg_count[d] += 1
            if len(neg_examples[d]) < 2:
                neg_examples[d].append(text[:80])

    all_dims = set(pos_count.keys()) | set(neg_count.keys())
    dim_stats = []
    for d in all_dims:
        p = pos_count[d]
        n = neg_count[d]
        t = p + n
        ratio = str(p) + '/' + str(t) if t > 0 else '0/0'
        net = p - n
        dim_stats.append((d, t, p, n, net, ratio))
    dim_stats.sort(key=lambda x: -x[1])

    header = '{:<12} {:>8} {:>8} {:>8} {:>8} {:>10}'.format('主题', '总提及', '正面', '负面', '净情感', '正/负比')
    print(header)
    print('-' * 60)
    for d, total, p, n, net, ratio in dim_stats:
        print('{:<12} {:>8} {:>8} {:>8} {:>+8} {:>10}'.format(d, total, p, n, net, ratio))

    print()
    print('--- 负面评论典型示例 ---')
    for d in sorted(neg_count.keys(), key=lambda x: -neg_count[x])[:5]:
        print('  [' + d + '] (neg=' + str(neg_count[d]) + ')')
        for ex in neg_examples[d][:1]:
            print('    "' + ex + '..."')

    return dict(pos_count), dict(neg_count), dict(neg_examples)

bx_pos, bx_neg, bx_neg_ex = analyze_hotel(bx, '佰翔')
zz_pos, zz_neg, zz_neg_ex = analyze_hotel(zz, '漳州')

# 关键事件
sep = '='*60
print(sep)
print('关键事件提取（极端评分）')
print(sep)

def extract_events(df, name, threshold=2.0, n=5):
    extreme = df[df['评分'] <= threshold]
    print()
    print('--- ' + name + ' 差评关键事件 (评分<=' + str(threshold) + ', 共' + str(len(extreme)) + '条) ---')
    for _, row in extreme.head(n).iterrows():
        text = str(row['评论内容'])[:150]
        print('  [' + str(row['评分']) + '分] ' + text + '...')
    print()

extract_events(bx, '佰翔', 2.0, 5)
extract_events(zz, '漳州', 2.0, 5)

# 竞品对比
sep = '='*60
print(sep)
print('佰翔 vs 漳州宾馆 - 维度对比')
print(sep)

all_dims = set(bx_pos.keys()) | set(bx_neg.keys()) | set(zz_pos.keys()) | set(zz_neg.keys())
header = '{:<12} {:>10} {:>10} {:>10} {:>10}'.format('维度', '佰翔正面', '佰翔负面', '漳州正面', '漳州负面')
print(header)
print('-' * 60)
for d in sorted(all_dims, key=lambda x: -(bx_pos.get(x,0)+bx_neg.get(x,0)+zz_pos.get(x,0)+zz_neg.get(x,0))):
    print('{:<12} {:>10} {:>10} {:>10} {:>10}'.format(d, bx_pos.get(d,0), bx_neg.get(d,0), zz_pos.get(d,0), zz_neg.get(d,0)))

# 顾客关注重点
sep = '='*60
print()
print(sep)
print('顾客关注重点分析（维度提及率）')
print(sep)

bx_total = bx['有效评论'].sum()
zz_total = zz['有效评论'].sum()
all_dims_sorted = sorted(all_dims, key=lambda x: -(bx_pos.get(x,0)+bx_neg.get(x,0)+zz_pos.get(x,0)+zz_neg.get(x,0)))

header = '{:<12} {:>14} {:>14} {:>10}'.format('维度', '佰翔提及率', '漳州提及率', '差异')
print(header)
print('-' * 60)
for d in all_dims_sorted[:10]:
    bx_rate = (bx_pos.get(d,0)+bx_neg.get(d,0)) / bx_total * 100
    zz_rate = (zz_pos.get(d,0)+zz_neg.get(d,0)) / zz_total * 100
    diff = bx_rate - zz_rate
    print('{:<12} {:>13.1f}% {:>13.1f}% {:>+9.1f}%'.format(d, bx_rate, zz_rate, diff))

# 好评驱动 vs 差评痛点
sep = '='*60
print()
print(sep)
print('好评驱动因素 vs 差评痛点')
print(sep)

def get_drivers(df):
    high = df[df['评分'] >= 4.5]
    pos_dims = Counter()
    for _, row in high.iterrows():
        text = str(row['评论内容'])
        p, _ = code_review(text)
        for d in p:
            pos_dims[d] += 1
    return pos_dims

def get_pain(df):
    low = df[df['评分'] < 3.0]
    neg_dims = Counter()
    for _, row in low.iterrows():
        text = str(row['评论内容'])
        _, n = code_review(text)
        for d in n:
            neg_dims[d] += 1
    return neg_dims

print()
print('佰翔 - 好评驱动因素 (4.5+分):')
bx_drivers = get_drivers(bx)
for d, c in bx_drivers.most_common(8):
    print('  ' + d + ': ' + str(c) + '次')

print()
print('佰翔 - 差评痛点 (<3分):')
bx_pain = get_pain(bx)
for d, c in bx_pain.most_common(8):
    print('  ' + d + ': ' + str(c) + '次')

print()
print('漳州 - 好评驱动因素 (4.5+分):')
zz_drivers = get_drivers(zz)
for d, c in zz_drivers.most_common(8):
    print('  ' + d + ': ' + str(c) + '次')

print()
print('漳州 - 差评痛点 (<3分):')
zz_pain = get_pain(zz)
for d, c in zz_pain.most_common(8):
    print('  ' + d + ': ' + str(c) + '次')

# 保存报告
print()
print('正在保存详细报告...')

with pd.ExcelWriter('酒店评论综合分析报告.xlsx', engine='openpyxl') as writer:
    # 维度对比
    rows = []
    for d in all_dims:
        rows.append({
            '维度': d,
            '佰翔正面': bx_pos.get(d,0),
            '佰翔负面': bx_neg.get(d,0),
            '佰翔净情感': bx_pos.get(d,0)-bx_neg.get(d,0),
            '漳州正面': zz_pos.get(d,0),
            '漳州负面': zz_neg.get(d,0),
            '漳州净情感': zz_pos.get(d,0)-zz_neg.get(d,0)
        })
    pd.DataFrame(rows).sort_values('佰翔正面', ascending=False).to_excel(writer, sheet_name='维度对比', index=False)

    # 佰翔差评详情
    bx_low = bx[bx['评分'] < 3.0][['评分','评论日期','评论内容']].copy()
    for idx, row in bx_low.iterrows():
        text = str(row['评论内容'])
        p, n = code_review(text)
        row['正面维度'] = '; '.join(p)
        row['负面维度'] = '; '.join(n)
    bx_low.to_excel(writer, sheet_name='佰翔差评详情', index=False)

    # 漳州差评详情
    zz_low = zz[zz['评分'] < 3.0][['评分','评论日期','评论内容']].copy()
    for idx, row in zz_low.iterrows():
        text = str(row['评论内容'])
        p, n = code_review(text)
        row['正面维度'] = '; '.join(p)
        row['负面维度'] = '; '.join(n)
    zz_low.to_excel(writer, sheet_name='漳州差评详情', index=False)

    # 佰翔好评示例
    bx_high = bx[bx['评分'] >= 4.5][['评分','评论日期','评论内容']].head(100).copy()
    for idx, row in bx_high.iterrows():
        text = str(row['评论内容'])
        p, n = code_review(text)
        row['正面维度'] = '; '.join(p)
        row['负面维度'] = '; '.join(n)
    bx_high.to_excel(writer, sheet_name='佰翔好评示例', index=False)

print()
print('报告已保存: 酒店评论综合分析报告.xlsx')
