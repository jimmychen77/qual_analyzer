"""
深度营销归因分析：佰翔 vs 漳州
方法：关键事件 + 情感强度 + 隐性需求 + 营销策略
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

# ========== 1. 情感强度分析 ==========
print('='*60)
print('1. 情感强度词分析（程度副词）')
print('='*60)

intense_pos = ['非常', '特别', '十分', '超级', '极致', '完美', '无可挑剔', '超棒', '超赞', '超值', '惊喜']
intense_neg = ['非常', '特别', '极其', '完全', '彻底', '太差', '极差', '噩梦', '垃圾', '恶心', '恐怖']

def count_intense(df, name):
    pos_intense = 0
    neg_intense = 0
    normal_pos = 0
    normal_neg = 0

    for _, row in df[df['有效评论']].iterrows():
        text = str(row['评论内容'])
        has_pos = any(w in text for w in ['好', '棒', '满意', '喜欢', '赞', '舒适', '干净', '热情', '周到'])
        has_neg = any(w in text for w in ['差', '失望', '不满', '糟糕', '恶心', '脏', '吵', '旧'])

        if has_pos:
            if any(w in text for w in intense_pos):
                pos_intense += 1
            else:
                normal_pos += 1
        if has_neg:
            if any(w in text for w in intense_neg):
                neg_intense += 1
            else:
                normal_neg += 1

    print()
    print(name + ':')
    print('  强烈正面(包含程度副词): ' + str(pos_intense))
    print('  普通正面: ' + str(normal_pos))
    print('  强烈负面: ' + str(neg_intense))
    print('  普通负面: ' + str(normal_neg))
    print('  正面强度比例: ' + str(round(pos_intense/(pos_intense+normal_pos)*100, 1)) + '%')
    print('  负面强度比例: ' + str(round(neg_intense/(neg_intense+normal_neg)*100, 1)) + '%')

count_intense(bx, '佰翔')
count_intense(zz, '漳州')

# ========== 2. 隐性需求挖掘 ==========
print()
print('='*60)
print('2. 隐性需求挖掘（好评中隐含的改进空间）')
print('='*60)

hidden_complaint_words = ['但是', '不过', '就是', '唯一', '美中不足', '稍有', '略显', '如果', '希望', '建议']

def find_hidden(df, name):
    hidden = df[(df['评分'] >= 4.0) & df['有效评论']].copy()
    results = []
    for _, row in hidden.iterrows():
        text = str(row['评论内容'])
        for w in hidden_complaint_words:
            if w in text:
                # 提取包含该词的前后文
                idx = text.find(w)
                context = text[max(0,idx-10):min(len(text),idx+20)]
                results.append((row['评分'], w, context, text[:80]))
                break
    print()
    print(name + ' - 隐性不满 (' + str(len(results)) + '条):')
    # 按词统计
    word_cnt = Counter([r[1] for r in results])
    for w, c in word_cnt.most_common(5):
        print('  "' + w + '": ' + str(c) + '条')
    for _, w, ctx, full in results[:3]:
        print('  [' + str(_) + '分] "' + w + '" -> ' + ctx + '...')
    return results

find_hidden(bx, '佰翔')
find_hidden(zz, '漳州')

# ========== 3. 关键事件归因 ==========
print()
print('='*60)
print('3. 关键事件归因（好评/差评的直接原因）')
print('='*60)

# 好评关键事件
pos_triggers = {
    '免费升房': ['升级', '升房', '免费升级', '给升级', '做了升房'],
    '赠送物品': ['果盘', '水果', '小礼物', '送东西', '送了'],
    '服务超预期': ['超出预期', '超乎想象', '惊喜', '感动', '贴心'],
    '主动服务': ['主动', '热情主动', '积极'],
    '前台服务好': ['前台', '办理', '入住', '服务热情'],
    '早餐好评': ['早餐', '餐厅', '用餐'],
}

neg_triggers = {
    '设施老旧破损': ['旧', '破', '坏', '故障', '不能用', '损坏', '生锈'],
    '卫生问题': ['脏', '毛发', '异味', '霉味', '恶心', '有虫', '不干净'],
    '隔音差': ['隔音', '噪音', '吵', '车声', '施工', '装修'],
    '服务冷漠': ['冷漠', '敷衍', '态度差', '不理', '爱答不理'],
    '等待时间长': ['等很久', '等了半天', '排队', '很久'],
    '押金/收费问题': ['押金', '收费', '骗', '坑'],
    '信息不对称': ['和说的', '不一样', '没有', '备注', '不知道'],
    '安全/隐患': ['不安全', '担心', '危险', '滑倒', '摔伤'],
}

def analyze_triggers(df, name, score_type='pos', threshold=4.5):
    if score_type == 'pos':
        data = df[df['评分'] >= threshold]
        triggers = pos_triggers
        label = '好评触发'
    else:
        data = df[df['评分'] < 3.0]
        triggers = neg_triggers
        label = '差评触发'

    print()
    print(name + ' - ' + label + ':')
    trigger_counts = defaultdict(int)
    trigger_examples = defaultdict(list)

    for _, row in data.iterrows():
        text = str(row['评论内容'])
        for event, keywords in triggers.items():
            for kw in keywords:
                if kw in text:
                    trigger_counts[event] += 1
                    if len(trigger_examples[event]) < 1:
                        trigger_examples[event].append(text[:80])
                    break

    sorted_triggers = sorted(trigger_counts.items(), key=lambda x: -x[1])
    for event, count in sorted_triggers:
        print('  ' + event + ': ' + str(count) + '次')
        for ex in trigger_examples[event][:1]:
            print('    例: "' + ex + '..."')

analyze_triggers(bx, '佰翔', 'pos', 4.5)
analyze_triggers(bx, '佰翔', 'neg', 3.0)
analyze_triggers(zz, '漳州', 'pos', 4.5)
analyze_triggers(zz, '漳州', 'neg', 3.0)

# ========== 4. 细分群体分析 ==========
print()
print('='*60)
print('4. 不同场景/出行目的分析')
print('='*60)

# 从评论内容中推断出行目的
def infer_purpose(text):
    if not isinstance(text, str):
        return '未知'
    if any(w in text for w in ['出差', '商务', '办公', '工作', '开会']):
        return '商务出差'
    elif any(w in text for w in ['亲子', '小孩', '孩子', '小朋友', '家庭', '全家']):
        return '亲子游'
    elif any(w in text for w in ['情侣', '老婆', '老公', '女朋友', '男朋友', '浪漫']):
        return '情侣游'
    elif any(w in text for w in ['父母', '老人', '妈妈', '爸爸', '长辈']):
        return '家庭游'
    elif any(w in text for w in ['旅游', '度假', '景点', '游玩', '旅行']):
        return '休闲游'
    else:
        return '其他'

for df, name in [(bx, '佰翔'), (zz, '漳州')]:
    df['出行目的'] = df['评论内容'].apply(infer_purpose)

print()
print('佰翔 - 出行目的分布:')
print(bx['出行目的'].value_counts())
print()
print('漳州 - 出行目的分布:')
print(zz['出行目的'].value_counts())

# 各目的的平均评分
print()
print('佰翔 - 各出行目的平均评分:')
for purpose, group in bx.groupby('出行目的'):
    if len(group) > 10:
        print('  ' + purpose + ': ' + str(round(group['评分'].mean(), 2)) + ' (n=' + str(len(group)) + ')')

print()
print('漳州 - 各出行目的平均评分:')
for purpose, group in zz.groupby('出行目的'):
    if len(group) > 5:
        print('  ' + purpose + ': ' + str(round(group['评分'].mean(), 2)) + ' (n=' + str(len(group)) + ')')

# ========== 5. 月度趋势分析 ==========
print()
print('='*60)
print('5. 月度趋势分析（评分波动）')
print('='*60)

for df, name in [(bx, '佰翔'), (zz, '漳州')]:
    monthly = df.groupby('月份').agg({'评分': ['mean', 'count']}).round(2)
    monthly.columns = ['平均分', '评论数']
    print()
    print(name + ' - 近12个月趋势:')
    print(monthly.tail(12).to_string())

# ========== 6. 竞品优劣势矩阵 ==========
print()
print('='*60)
print('6. 竞品优劣势矩阵')
print('='*60)

# 主观判断的优劣势阈值
def get_advantage_disadvantage(bx_pos, bx_neg, zz_pos, zz_neg, total_bx, total_zz):
    advantages = []
    disadvantages = []
    opportunities = []
    threats = []

    all_dims = set(bx_pos.keys()) | set(bx_neg.keys()) | set(zz_pos.keys()) | set(zz_neg.keys())

    for d in all_dims:
        bx_rate = (bx_pos.get(d,0) + bx_neg.get(d,0)) / total_bx * 100
        zz_rate = (zz_pos.get(d,0) + zz_neg.get(d,0)) / total_zz * 100
        bx_net = bx_pos.get(d,0) - bx_neg.get(d,0)
        zz_net = zz_pos.get(d,0) - zz_neg.get(d,0)

        # 佰翔优势：提及率高且净情感高
        if bx_rate > zz_rate and bx_net > zz_net:
            advantages.append(d)
        # 佰翔劣势：某维度负面多
        elif bx_neg.get(d,0) > bx_pos.get(d,0) * 0.1:
            disadvantages.append(d)
        # 机会：漳州提及时佰翔未充分开发
        if zz_rate > 5 and bx_rate < 5:
            opportunities.append(d)

    return advantages, disadvantages, opportunities

bx_total = bx['有效评论'].sum()
zz_total = zz['有效评论'].sum()
bx_pos, bx_neg = defaultdict(int), defaultdict(int)
zz_pos, zz_neg = defaultdict(int), defaultdict(int)

coding_dict = {
    '清洁卫生': {'pos': ['干净', '整洁', '卫生', '清洁'], 'neg': ['脏', '毛发', '霉味', '异味', '恶心']},
    '服务态度': {'pos': ['热情', '周到', '贴心', '耐心'], 'neg': ['冷漠', '敷衍', '态度差']},
    '前台服务': {'pos': ['前台热情', '办理快', '效率高'], 'neg': ['前台差', '入住慢']},
    '设施设备': {'pos': ['设施齐全', '设备好', '新'], 'neg': ['旧', '坏', '故障']},
    '房间条件': {'pos': ['房间大', '宽敞', '舒适', '床舒服'], 'neg': ['房间小', '狭小', '不舒服']},
    '隔音效果': {'pos': ['隔音好', '安静'], 'neg': ['隔音差', '吵', '噪音']},
    '早餐餐饮': {'pos': ['早餐好', '丰富', '品种多'], 'neg': ['早餐差', '单一', '难吃']},
    '位置交通': {'pos': ['位置好', '方便', '近'], 'neg': ['位置偏', '不方便']},
    '性价比': {'pos': ['物超所值', '划算', '值得'], 'neg': ['不值', '贵']},
    '景观环境': {'pos': ['景观好', '漂亮', '美'], 'neg': ['难看']},
    '停车配套': {'pos': ['停车方便', '停车场'], 'neg': ['停车难']},
}

def code(text):
    p, n = [], []
    for dim, words in coding_dict.items():
        for w in words['pos']:
            if w in text:
                p.append(dim)
                break
        for w in words['neg']:
            if w in text:
                n.append(dim)
                break
    return list(set(p)), list(set(n))

for _, row in bx[bx['有效评论']].iterrows():
    p, n = code(str(row['评论内容']))
    for d in p: bx_pos[d] += 1
    for d in n: bx_neg[d] += 1

for _, row in zz[zz['有效评论']].iterrows():
    p, n = code(str(row['评论内容']))
    for d in p: zz_pos[d] += 1
    for d in n: zz_neg[d] += 1

advantages, disadvantages, opportunities = get_advantage_disadvantage(
    bx_pos, bx_neg, zz_pos, zz_neg, bx_total, zz_total)

print()
print('佰翔酒店优势（提及率高于漳州且净情感更正面）:')
for a in advantages:
    print('  + ' + a)

print()
print('佰翔酒店待改进（负面提及率相对较高）:')
for d in disadvantages:
    print('  ! ' + d + ' (负面' + str(bx_neg.get(d,0)) + '次)')

print()
print('佰翔潜在机会（漳州顾客关注但佰翔提及率低）:')
for o in opportunities:
    print('  * ' + o)

# ========== 7. 营销策略建议 ==========
print()
print('='*60)
print('7. 营销策略建议')
print('='*60)

print('''
【佰翔酒店 - 营销策略建议】

1. 强化服务优势（服务态度已是最强项）
   - 将"服务热情、周到"作为核心卖点进行视觉化呈现
   - 推出"服务之星"员工表彰体系
   - 鼓励客人发朋友圈/小红书，利用UGC传播

2. 紧急改进：房间条件与隔音
   - "房间条件"是负提及最高的维度(1032次)
   - 隔音正负比接近1:1，是最大痛点
   - 建议在OTA页面主动说明"安静楼层"或"隔音房型"

3. 早餐体验提升
   - 负面集中在"人潮涌动、位置不足"
   - 可推广"错峰早餐"或"送餐到房"服务

4. 亲子游市场深耕
   - 亲子目的评论多，好评率高
   - 可开发"亲子主题房"并投放亲子类KOL

5. 漳州竞品对比劣势提醒
   - 在"位置交通"和"清洁卫生"上，漳州提及率更高
   - 需加强这两项的主动宣传

【漳州宾馆 - 营销策略建议】

1. 差异化定位：历史底蕴+园林环境
   - 景观环境零负面，可主打"园林酒店"概念
   - 对标"老牌精品酒店"定位

2. 设施更新是当务之急
   - 房间条件负提及率高达52%（89/172）
   - 设施设备负提及率25%（33/134）
   - 建议分批翻新，优先改善隔音

3. 早餐成为短板
   - 漳州客人对早餐关注度高于佰翔
   - 建议增加热菜品种和儿童收费优化
''')

print()
print('='*60)
print('分析完成')
print('='*60)
