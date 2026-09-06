# =============================================
# 《明日方舟》玩家问卷 - 描述性统计与交叉分析
# =============================================
# 【文件说明】
#   本脚本对清洗后的问卷数据进行全面的描述性统计和交叉分析，
#   输出终端报告、Excel 汇总表和可视化图表。
#
# 【分析框架】
#   第一部分：描述性统计 —— 玩家画像、满意度、痛点、付费、NPS 等
#   第二部分：竞品深度分析 —— 从"玩过游戏"列提取竞品洞察
#   第三部分：H1-H4 假设验证 —— 事先有猜想，用数据验证
#   第四部分：C1-C5 探索性分析 —— 在数据中寻找未知规律
#   第五部分：导出 Excel —— 所有表格数据汇总
#   第六部分：生成图表 —— 可视化输出
#
# 【方法说明】
#   采用大样本统计方法（均值、百分比、交叉表等），
#   当前样本量较小（n=28），统计推断效力有限，结果仅供探索性参考。
# =============================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
from math import pi
from collections import Counter
import os

warnings.filterwarnings('ignore')

# ------------------------------
# 1. 全局设置
# ------------------------------
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False


# ------------------------------
# 2. 读取数据
# ------------------------------
print("=" * 80)
print("《明日方舟》玩家问卷 - 描述性统计与交叉分析")
print("=" * 80)
print("⚠️  方法说明：以下分析采用大样本统计方法（均值、百分比、交叉表等）")
print("⚠️  当前样本量较小，统计推断效力有限，结果仅供探索性参考。")
print("=" * 80)

df = pd.read_excel('清洗后问卷数据.xlsx')
total_n = len(df)
print(f"\n📌 当前样本量：{total_n} 份")

# 数值列缺失值填充 0
numeric_cols = df.select_dtypes(include=[np.number]).columns
df[numeric_cols] = df[numeric_cols].fillna(0)

# ------------------------------
# 3. 定义关键列名（便于后续引用）
# ------------------------------
NPS_COL = '18、您有多大可能性将《明日方舟》推荐给朋友？'
PAY_COL = '13、您在《明日方舟》中的累计付费金额大约是多少？'
WILLING_COL = '20、假设《明日方舟》计划推出一个全新的【常驻联机合作玩法】，请问您的参与意愿如何？'

# 玩家类型列（Q7 拆分后）
TYPE_COLS = [c for c in df.columns if '党' in c and (
    '收集' in c or '剧情' in c or '强度' in c or '休闲' in c or 'XP' in c)]

# 痛点列（Q10 拆分后）
PAIN_COLS = ['抽卡概率低', '资源获取速度慢', '关卡难度过高', '活动内容重复',
             '干员平衡性问题', '社交功能薄弱', '没有明显不满']
PAIN_COLS = [p for p in PAIN_COLS if p in df.columns]

# 满意度列
SAT_COLS = ['满意_美术', '满意_剧情', '满意_关卡', '满意_音乐', '满意_整体']


# ------------------------------
# 4. 辅助函数
# ------------------------------
def print_table(title, series):
    """打印频次表格"""
    print(f"\n【{title}】")
    table = pd.DataFrame({
        '选项': series.index,
        '人数': series.values,
        '占比': [f"{v / total_n * 100:.1f}%" for v in series.values]
    })
    print(table.to_string(index=False))


# =============================================
# 第一部分：描述性统计
# =============================================

# ------------------------------
# 4.1 玩家画像
# ------------------------------
print("\n" + "=" * 80)
print("一、玩家画像")
print("=" * 80)

for col in ['1、您的性别是？', '2、您的年龄段是', '3、您目前的职业是',
            '4、您玩《明日方舟》多长时间了？', '5、您一般每天花多少时间在《明日方舟》上？']:
    if col in df.columns:
        print_table(col, df[col].value_counts())
        if total_n < 30:
            print(f"  ⚠️ 样本量小（n={total_n}），该占比仅供参考")

# ------------------------------
# 4.2 时间分配
# ------------------------------
print("\n" + "=" * 80)
print("二、时间分配")
print("=" * 80)

time_cols = [c for c in df.columns if any(kw in c for kw in
                                           ['清日常', '集成战略', '卫戍协议', '观看剧情', '研究关卡', '基建'])]
time_data = []
for c in time_cols:
    cnt = df[c].sum()
    if cnt > 0:
        time_data.append({'选项': c, '人数': cnt, '占比': f"{cnt / total_n * 100:.1f}%"})
if time_data:
    print(pd.DataFrame(time_data).to_string(index=False))
    print(f"  ⚠️ 样本量小（n={total_n}），该占比仅供参考")

# ------------------------------
# 4.3 玩家类型
# ------------------------------
print("\n" + "=" * 80)
print("三、玩家类型")
print("=" * 80)

type_data = []
for c in TYPE_COLS:
    cnt = df[c].sum()
    if cnt > 0:
        type_data.append({'类型': c, '人数': cnt, '占比': f"{cnt / total_n * 100:.1f}%"})
if type_data:
    print(pd.DataFrame(type_data).to_string(index=False))

# ------------------------------
# 4.4 满意度
# ------------------------------
print("\n" + "=" * 80)
print("四、满意度")
print("=" * 80)

sat_data = []
for c in SAT_COLS:
    if c in df.columns:
        sat_data.append({
            '维度': c.replace('满意_', ''),
            '均值': f"{df[c].mean():.2f}",
            '标准差': f"{df[c].std():.2f}",
            '中位数': f"{df[c].median():.0f}",
            '最小值': f"{df[c].min():.0f}",
            '最大值': f"{df[c].max():.0f}"
        })
if sat_data:
    print(pd.DataFrame(sat_data).to_string(index=False))
    print(f"  ⚠️ 小样本（n={total_n}），标准差参考价值有限")

# 满意度分布
if '满意_整体' in df.columns:
    print("\n【满意度整体分布】")
    dist_data = []
    for score, cnt in df['满意_整体'].value_counts().sort_index().items():
        dist_data.append({'评分': score, '人数': cnt, '占比': f"{cnt / total_n * 100:.1f}%"})
    print(pd.DataFrame(dist_data).to_string(index=False))

# ------------------------------
# 4.5 痛点
# ------------------------------
print("\n" + "=" * 80)
print("五、痛点")
print("=" * 80)

pain_data = []
for p in PAIN_COLS:
    cnt = df[p].sum()
    pain_data.append({'痛点': p, '人数': cnt, '提及率': f"{cnt / total_n * 100:.1f}%"})
if pain_data:
    print(pd.DataFrame(sorted(pain_data, key=lambda x: int(x['人数']), reverse=True)).to_string(index=False))

# ------------------------------
# 4.6 付费
# ------------------------------
print("\n" + "=" * 80)
print("六、付费")
print("=" * 80)

if PAY_COL in df.columns:
    pay_data = []
    for k, v in df[PAY_COL].value_counts().items():
        pay_data.append({'金额区间': k, '人数': v, '占比': f"{v / total_n * 100:.1f}%"})
    print(pd.DataFrame(pay_data).to_string(index=False))

# 付费动机
print("\n【付费动机】")
motivation_cols = ['为喜爱的角色/厨力', '为提升游戏强度/通关效率', '为好看的皮肤/外观',
                   '为支持游戏/鹰角网络', '月卡/礼包性价比高']
motiv_data = []
for m in motivation_cols:
    if m in df.columns:
        cnt = df[m].sum()
        motiv_data.append({'动机': m, '人数': cnt, '提及率': f"{cnt / total_n * 100:.1f}%"})
if motiv_data:
    print(pd.DataFrame(motiv_data).to_string(index=False))

# 性价比感受
if '16、您对目前游戏内礼包/月卡的性价比感受如何？' in df.columns:
    print("\n【性价比感受】")
    value_col = '16、您对目前游戏内礼包/月卡的性价比感受如何？'
    val_data = []
    for k, v in df[value_col].value_counts().items():
        val_data.append({'感受': k, '人数': v, '占比': f"{v / total_n * 100:.1f}%"})
    print(pd.DataFrame(val_data).to_string(index=False))

# ------------------------------
# 4.7 集成战略（肉鸽）
# ------------------------------
print("\n" + "=" * 80)
print("七、集成战略（肉鸽）")
print("=" * 80)

if '肉鸽_参与度' in df.columns:
    print(f"参与度均值: {df['肉鸽_参与度'].mean():.2f} | 满意度均值: {df['肉鸽_满意度'].mean():.2f}")
    print("\n【参与度分布】")
    part_data = []
    for score, cnt in df['肉鸽_参与度'].value_counts().sort_index().items():
        part_data.append({'评分': score, '人数': cnt, '占比': f"{cnt / total_n * 100:.1f}%"})
    print(pd.DataFrame(part_data).to_string(index=False))

# ------------------------------
# 4.8 活动评价
# ------------------------------
if '12、您对《明日方舟》近期活动/新内容的整体评价是？' in df.columns:
    print("\n" + "=" * 80)
    print("八、近期活动评价")
    print("=" * 80)
    activity_col = '12、您对《明日方舟》近期活动/新内容的整体评价是？'
    act_data = []
    for k, v in df[activity_col].value_counts().items():
        act_data.append({'评价': k, '人数': v, '占比': f"{v / total_n * 100:.1f}%"})
    print(pd.DataFrame(act_data).to_string(index=False))

# ------------------------------
# 4.9 NPS 净推荐值
# ------------------------------
print("\n" + "=" * 80)
print("九、NPS净推荐值")
print("=" * 80)

if NPS_COL in df.columns:
    nps_vals = df[NPS_COL].dropna()
    recommenders = (nps_vals >= 9).sum()
    passives = nps_vals[nps_vals.between(7, 8)].count()
    detractors = (nps_vals <= 6).sum()
    nps_score = (recommenders - detractors) / len(nps_vals) * 100 if len(nps_vals) > 0 else 0

    nps_data = [
        {'指标': '平均分', '数值': f"{nps_vals.mean():.2f}"},
        {'指标': '中位数', '数值': f"{nps_vals.median():.0f}"},
        {'指标': '推荐者(9-10)', '数值': f"{recommenders}人 ({recommenders / len(nps_vals) * 100:.1f}%)"},
        {'指标': '被动者(7-8)', '数值': f"{passives}人 ({passives / len(nps_vals) * 100:.1f}%)"},
        {'指标': '贬损者(0-6)', '数值': f"{detractors}人 ({detractors / len(nps_vals) * 100:.1f}%)"},
        {'指标': 'NPS', '数值': f"{nps_score:.1f}"}
    ]
    print(pd.DataFrame(nps_data).to_string(index=False))
    print(f"  ⚠️ 小样本（n={len(nps_vals)}），NPS值波动较大，仅供参考")

# ------------------------------
# 4.10 联机玩法
# ------------------------------
print("\n" + "=" * 80)
print("十、联机玩法")
print("=" * 80)

if WILLING_COL in df.columns:
    will_data = []
    for k, v in df[WILLING_COL].value_counts().items():
        will_data.append({'意愿': k, '人数': v, '占比': f"{v / total_n * 100:.1f}%"})
    print(pd.DataFrame(will_data).to_string(index=False))

if '21、您最看重该玩法的哪个方面？' in df.columns:
    print("\n【最看重方面】")
    focus_col = '21、您最看重该玩法的哪个方面？'
    focus_data = []
    for k, v in df[focus_col].value_counts().items():
        focus_data.append({'方面': k, '人数': v, '占比': f"{v / total_n * 100:.1f}%"})
    print(pd.DataFrame(focus_data).to_string(index=False))

# ------------------------------
# 4.11 其他内容列汇总
# ------------------------------
print("\n" + "=" * 80)
print("十一、其他内容列汇总")
print("=" * 80)

other_cols = [c for c in df.columns if c.endswith('_其他内容')]
if other_cols:
    for col in other_cols:
        non_empty = df[df[col].notna() & (df[col] != '')]
        if len(non_empty) > 0:
            print(f"\n【{col}】共 {len(non_empty)} 条：")
            for _, row in non_empty.iterrows():
                print(f"  用户{row.name}: {row[col]}")
else:
    print("  ℹ️ 无其他内容列")


# =============================================
# 第二部分：竞品深度分析
# =============================================
print("\n" + "=" * 80)
print("十二、竞品深度分析")
print("=" * 80)

if '玩过游戏' in df.columns:
    # 提取所有游戏名称
    game_list = []
    for val in df['玩过游戏'].dropna():
        if val and isinstance(val, str):
            for g in val.split(','):
                g = g.strip()
                if g:
                    game_list.append(g)

    if game_list:
        game_counts = Counter(game_list)

        # 12.1 TOP 游戏排名
        print("\n【玩家提及的其他游戏 TOP10】")
        game_data = []
        for game, count in game_counts.most_common(10):
            game_data.append({
                '游戏': game,
                '人数': count,
                '提及率': f"{count / total_n * 100:.1f}%"
            })
        print(pd.DataFrame(game_data).to_string(index=False))

        # 12.2 竞品分类统计
        competitor_keywords = {
            '直接竞品（开放世界）': ['原神', '鸣潮', '绝区零'],
            '同赛道（二次元RPG）': ['崩坏星穹铁道', '碧蓝航线', 'FGO', '重返未来1999'],
            '衍生作品': ['明日方舟终末地', '终末地'],
            '其他二游': ['边狱巴士', '少前2', '异环', '忘却前夜'],
            '非二游': ['英雄联盟', '巫师三', '战争雷霆', '音乐游戏', 'p5r']
        }

        print("\n【竞品分类统计】")
        category_counts = {k: 0 for k in competitor_keywords.keys()}
        for game, count in game_counts.items():
            for cat, keywords in competitor_keywords.items():
                if game in keywords:
                    category_counts[cat] += count
                    break

        for cat, count in category_counts.items():
            if count > 0:
                print(f"  {cat}: {count}次提及")

        # 12.3 玩家游戏数量分布
        game_count_per_user = df['玩过游戏'].apply(
            lambda x: len([g for g in str(x).split(',') if g.strip()]) if x and isinstance(x, str) else 0
        )
        print("\n【玩家提及游戏数量分布】")
        for cnt in sorted(game_count_per_user[game_count_per_user > 0].value_counts().items()):
            print(f"  提及{cnt[0]}款游戏: {cnt[1]}人")

        # 12.4 玩竞品 vs 不玩竞品的满意度对比
        has_competitor = df['玩过游戏'].apply(
            lambda x: any(g in str(x) for g in ['原神', '鸣潮', '绝区零', '崩坏星穹铁道']) if x else False
        )
        if '满意_整体' in df.columns:
            has_comp_sat = df[has_competitor]['满意_整体'].mean()
            no_comp_sat = df[~has_competitor]['满意_整体'].mean()
            print(f"\n【玩竞品 vs 不玩竞品 满意度对比】")
            print(f"  玩竞品玩家满意度: {has_comp_sat:.2f} (n={has_competitor.sum()})")
            print(f"  不玩竞品玩家满意度: {no_comp_sat:.2f} (n={(~has_competitor).sum()})")

        # 12.5 玩竞品 vs 不玩竞品的付费对比
        if PAY_COL in df.columns:
            pay_high = '5000元以上'
            has_comp_high = df[has_competitor & (df[PAY_COL] == pay_high)]
            no_comp_high = df[(~has_competitor) & (df[PAY_COL] == pay_high)]
            has_comp_total = has_competitor.sum()
            no_comp_total = (~has_competitor).sum()
            if has_comp_total > 0 and no_comp_total > 0:
                print(f"\n【玩竞品 vs 不玩竞品 付费对比】")
                print(f"  玩竞品玩家高付费率: {len(has_comp_high)}/{has_comp_total} ({len(has_comp_high) / has_comp_total * 100:.1f}%)")
                print(f"  不玩竞品玩家高付费率: {len(no_comp_high)}/{no_comp_total} ({len(no_comp_high) / no_comp_total * 100:.1f}%)")

        # 12.6 竞品接触数量与满意度的关系
        if '满意_整体' in df.columns:
            print("\n【竞品接触数量 × 满意度】")
            df['游戏提及数'] = game_count_per_user
            sat_by_count = df.groupby('游戏提及数')['满意_整体'].mean()
            for k, v in sat_by_count.items():
                if k > 0:
                    print(f"  提及{k}款游戏: {v:.2f} (n={len(df[df['游戏提及数'] == k])})")

    else:
        print("  ℹ️ 无游戏提及数据")
else:
    print("  ⚠️ 未找到'玩过游戏'列，请重新运行 clean_data.py")


# =============================================
# 第三部分：H1-H4 假设验证
# =============================================
print("\n" + "=" * 80)
print("十三、H1-H4 假设验证（验证性分析）")
print("=" * 80)
print("   H = Hypothesis，事先有猜想，用数据验证")
print("   以下分析回答：我猜可能是这样，数据怎么说？")
print("-" * 80)
print(f"⚠️  当前样本量 {total_n} 份，分组后多数子组 n<10，统计效力不足。")
print("    以下结果仅为探索性参考，正式结论需扩大样本后验证。")
print("=" * 80)


# ---------- H1：玩家类型 × 付费金额 ----------
print("\n【H1】玩家类型 × 付费金额")
print("  假设：XP 党 / 收集党 可能是付费主力")
print("-" * 40)

if TYPE_COLS and PAY_COL in df.columns:
    h1_results = {}
    for t in TYPE_COLS:
        type_df = df[df[t] == 1]
        if len(type_df) > 0:
            pay_dist = type_df[PAY_COL].value_counts(normalize=True) * 100
            h1_results[t] = pay_dist

    if h1_results:
        h1_df = pd.DataFrame(h1_results).fillna(0)
        print("\n  各玩家类型付费分布（%）：")
        print(h1_df.round(1).to_string())

        print("\n  各类型样本量：")
        for t in TYPE_COLS:
            n = len(df[df[t] == 1])
            print(f"    {t[:15]}...: n={n}")

        # ✅ 生成图表：竖着的柱状图，文字横着
        fig, ax = plt.subplots(figsize=(12, 7))

        # 简化类型名称（只保留"X党"）
        short_labels = []
        for label in h1_df.columns:
            if '—' in label:
                short = label.split('—')[0].strip()
            elif '党' in label:
                short = label[:4]
            else:
                short = label[:6]
            short_labels.append(short)

        # 竖着绘图（kind='bar'）
        h1_df.T.plot(kind='bar', stacked=True, ax=ax, colormap='viridis', width=0.7)

        # 关键：x轴标签横着放（rotation=0），居中
        ax.set_xticklabels(short_labels, rotation=0, ha='center', fontsize=11)

        ax.set_title('各玩家类型付费分布', fontsize=14, pad=20)
        ax.set_xlabel('玩家类型', fontsize=12, labelpad=10)
        ax.set_ylabel('占比（%）', fontsize=12, labelpad=10)

        # 图例放在右侧外部
        ax.legend(title='付费金额', bbox_to_anchor=(1.02, 1), loc='upper left', fontsize=9)

        # 调整边距，给图例留空间
        plt.subplots_adjust(right=0.78, bottom=0.15, top=0.92)
        plt.savefig('cross_type_pay.png', dpi=300, bbox_inches='tight')
        print("  ✅ 已保存：cross_type_pay.png")

# ---------- H2：痛点 × NPS ----------
print("\n" + "-" * 40)
print("【H2】痛点 × NPS")
print("  假设：抽卡体验可能是最影响推荐意愿的痛点")
print("-" * 40)

if PAIN_COLS and NPS_COL in df.columns:
    h2_data = []
    for p in PAIN_COLS:
        has_pain = df[df[p] == 1]
        no_pain = df[df[p] == 0]
        if len(has_pain) >= 3 and len(no_pain) >= 3:
            h2_data.append({
                '痛点': p,
                '有此痛点NPS': has_pain[NPS_COL].mean(),
                '无此痛点NPS': no_pain[NPS_COL].mean(),
                '差距': no_pain[NPS_COL].mean() - has_pain[NPS_COL].mean(),
                '有此痛点人数': len(has_pain)
            })

    if h2_data:
        h2_df = pd.DataFrame(h2_data).sort_values('差距', ascending=False)
        print("\n  各痛点对NPS的影响：")
        print(h2_df.round(2).to_string(index=False))
    else:
        print("  ⚠️ 样本量不足（各痛点的有/无分组 < 3），跳过 H2")
else:
    print("  ⚠️ 数据不足，跳过 H2")


# ---------- H3：联机意愿 × 满意度 ----------
print("\n" + "-" * 40)
print("【H3】联机意愿 × 整体满意度")
print("  假设：高满意度玩家可能更期待联机玩法")
print("-" * 40)

if WILLING_COL in df.columns:
    print("\n  联机意愿分布：")
    will_dist = df[WILLING_COL].value_counts()
    for k, v in will_dist.items():
        print(f"    {k}: {v}人 ({v / total_n * 100:.1f}%)")

    if '满意_整体' in df.columns:
        print("\n  联机意愿 × 整体满意度：")
        will_sat = df.groupby(WILLING_COL)['满意_整体'].agg(['mean', 'count'])
        for idx, row in will_sat.iterrows():
            if row['count'] >= 3:
                print(f"    {idx}: 均值{row['mean']:.2f} (n={row['count']})")
            else:
                print(f"    {idx}: 均值{row['mean']:.2f} (n={row['count']}) ⚠️ 样本量<3")
else:
    print("  ⚠️ 数据不足，跳过 H3")


# ---------- H4：玩家类型 × 行为验证 ----------
print("\n" + "-" * 40)
print("【H4】玩家类型 × 行为验证")
print("  假设：自述'剧情党'的玩家可能真的花更多时间看剧情")
print("-" * 40)

if TYPE_COLS:
    behavior_cols = [c for c in df.columns if any(kw in c for kw in
                                                   ['清日常', '集成战略', '卫戍协议', '观看剧情', '研究关卡', '基建'])]

    if behavior_cols:
        print("\n  各类型行为偏好矩阵（%）：")
        behavior_matrix = {}
        for t in TYPE_COLS:
            type_df = df[df[t] == 1]
            type_n = len(type_df)
            if type_n >= 3:
                if '—' in t:
                    label = t.split('—')[0].strip()
                else:
                    label = t[:8]
                behavior_matrix[label] = {}
                for b in behavior_cols[:6]:
                    behavior_matrix[label][b] = type_df[b].sum() / type_n * 100

        if behavior_matrix:
            h4_matrix = pd.DataFrame(behavior_matrix).fillna(0)
            print(h4_matrix.round(1).to_string())

            # ✅ 生成热力图：所有文字横着
            fig, ax = plt.subplots(figsize=(12, 8))

            sns.heatmap(h4_matrix, annot=True, fmt='.1f', cmap='Blues',
                        cbar_kws={'label': '参与率（%）'}, ax=ax,
                        annot_kws={'fontsize': 10})

            plt.title('各玩家类型行为偏好矩阵', fontsize=14, pad=20)
            plt.xlabel('玩家类型', fontsize=12, labelpad=10)
            plt.ylabel('行为', fontsize=12, labelpad=10)

            # 所有标签横着放：y轴标签 rotation=0，x轴标签也 rotation=0
            ax.set_yticklabels(ax.get_yticklabels(), rotation=0, ha='right', fontsize=10)
            ax.set_xticklabels(ax.get_xticklabels(), rotation=0, ha='center', fontsize=10)

            plt.subplots_adjust(bottom=0.2, left=0.2, top=0.92, right=0.92)
            plt.savefig('cross_type_behavior.png', dpi=300, bbox_inches='tight')
            print("  ✅ 已保存：cross_type_behavior.png")
        else:
            print("  ⚠️ 各类型样本量均 < 3，跳过 H4 图表生成")
    else:
        print("  ⚠️ 未找到行为列")
else:
    print("  ⚠️ 数据不足，跳过 H4")


print("\n" + "=" * 80)
print(f"📌 H1-H4 汇总：当前分析基于 {total_n} 份有效问卷。")
print("   交叉分析结果因样本量不足，不作为最终结论。")
print("   建议后续扩大样本量（目标 n≥100）后重新运行本脚本。")
print("=" * 80)


# =============================================
# 第四部分：C1-C5 探索性分析（贴合 JD）
# =============================================
print("\n" + "=" * 80)
print("十四、C1-C5 探索性分析（探索性分析）")
print("=" * 80)
print("   C = Cross，事先不知道会怎样，用数据发现规律")
print("   以下分析回应该岗位的核心职责：")
print("   · 协助产品调优，改善游戏体验")
print("   · 构建玩家画像，深入挖掘用户价值")
print("   · 针对商业化、玩家生态开展专项分析")
print("-" * 80)
print("⚠️  小样本提示：以下分析仅展示可用的数据，样本量不足时自动跳过。")
print("=" * 80)


# ---------- C1：玩家类型 × 满意度 ----------
print("\n【C1】玩家类型 × 满意度")
print("  业务问题：不同类型的玩家，满意度短板分别是什么？")
print("-" * 40)

if TYPE_COLS and all(c in df.columns for c in SAT_COLS):
    c1_data = []
    for t in TYPE_COLS:
        type_n = len(df[df[t] == 1])
        if type_n >= 3:
            row = {'玩家类型': t.replace(' — 喜欢...', '').replace('党', '党')[:10], '人数': type_n}
            for sat in SAT_COLS:
                row[sat.replace('满意_', '')] = df[df[t] == 1][sat].mean()
            c1_data.append(row)

    if c1_data:
        c1_df = pd.DataFrame(c1_data)
        print(c1_df.round(2).to_string(index=False))
        print("\n  💡 业务含义：找出各类型玩家的满意度低谷，指导版本调优方向")

        # 生成热力图
        plt.figure(figsize=(10, 6))
        heat_data = c1_df.set_index('玩家类型')[['美术', '剧情', '关卡', '音乐', '整体']]
        sns.heatmap(heat_data, annot=True, fmt='.2f', cmap='RdYlGn', center=3.5)
        plt.title('各玩家类型满意度热力图', fontsize=14)
        plt.tight_layout()
        plt.savefig('insight_type_satisfaction.png', dpi=300)
        print("  ✅ 已保存：insight_type_satisfaction.png")
    else:
        print("  ⚠️ 各类型样本量均 < 3，跳过 C1")
else:
    print("  ⚠️ 数据不足，跳过 C1")


# ---------- C2：付费 × 满意度 ----------
print("\n" + "-" * 40)
print("【C2】付费 × 满意度")
print("  业务问题：核心付费用户的真实体验如何？")
print("-" * 40)

if PAY_COL in df.columns and '满意_整体' in df.columns:
    pay_groups = {
        '零氪': ['0元（零氪党）'],
        '轻氪': ['1-100元', '101-500元'],
        '中氪': ['501-2000元', '2001-5000元'],
        '重氪': ['5000元以上']
    }

    c2_data = []
    for group_name, pay_levels in pay_groups.items():
        group_df = df[df[PAY_COL].isin(pay_levels)]
        if len(group_df) >= 3:
            row = {'付费等级': group_name, '人数': len(group_df)}
            for sat in SAT_COLS:
                row[sat.replace('满意_', '')] = group_df[sat].mean()
            c2_data.append(row)

    if c2_data:
        c2_df = pd.DataFrame(c2_data)
        print(c2_df.round(2).to_string(index=False))
        print("\n  💡 业务含义：重氪玩家的满意度痛点就是商业化优化的第一优先级")

        # 生成分组柱状图
        fig, ax = plt.subplots(figsize=(10, 6))
        c2_df.set_index('付费等级')[['美术', '剧情', '关卡', '音乐', '整体']].plot(kind='bar', ax=ax)
        ax.set_title('各付费等级满意度对比', fontsize=14)
        ax.set_ylabel('满意度均值')
        ax.legend(title='维度', bbox_to_anchor=(1.05, 1))
        plt.tight_layout()
        plt.savefig('insight_pay_satisfaction.png', dpi=300)
        print("  ✅ 已保存：insight_pay_satisfaction.png")
    else:
        print("  ⚠️ 各付费等级样本量均 < 3，跳过 C2")
else:
    print("  ⚠️ 数据不足，跳过 C2")


# ---------- C3：玩家类型 × 痛点 ----------
print("\n" + "-" * 40)
print("【C3】玩家类型 × 痛点")
print("  业务问题：不同玩家群体，抱怨的东西是一样的吗？")
print("-" * 40)

if TYPE_COLS and PAIN_COLS:
    c3_data = []
    for t in TYPE_COLS:
        type_n = len(df[df[t] == 1])
        if type_n >= 3:
            row = {'玩家类型': t.replace(' — 喜欢...', '').replace('党', '党')[:10], '人数': type_n}
            top_pains = []
            for p in PAIN_COLS:
                rate = df[df[t] == 1][p].sum() / type_n * 100
                row[p] = f"{rate:.0f}%"
                if rate > 30:
                    top_pains.append(p)
            row['TOP痛点'] = '、'.join(top_pains) if top_pains else '无明显痛点'
            c3_data.append(row)

    if c3_data:
        c3_df = pd.DataFrame(c3_data)
        print(c3_df.to_string(index=False))
        print("\n  💡 业务含义：不同群体的核心痛点不同，运营策略应差异化")
    else:
        print("  ⚠️ 各类型样本量均 < 3，跳过 C3")
else:
    print("  ⚠️ 数据不足，跳过 C3")


# ---------- C4：痛点 × 付费 ----------
print("\n" + "-" * 40)
print("【C4】痛点 × 付费")
print("  业务问题：高付费玩家的痛点和零氪玩家有什么不同？")
print("-" * 40)

if PAIN_COLS and PAY_COL in df.columns:
    high_pay = df[df[PAY_COL].isin(['2001-5000元', '5000元以上'])]
    low_pay = df[df[PAY_COL].isin(['0元（零氪党）', '1-100元'])]

    if len(high_pay) >= 3 and len(low_pay) >= 3:
        print("\n  高付费玩家（n={}）vs 零氪/轻氪玩家（n={}）".format(len(high_pay), len(low_pay)))
        c4_data = []
        for p in PAIN_COLS:
            high_rate = high_pay[p].sum() / len(high_pay) * 100
            low_rate = low_pay[p].sum() / len(low_pay) * 100
            diff = high_rate - low_rate
            c4_data.append({
                '痛点': p,
                '高付费提及率': f"{high_rate:.0f}%",
                '低付费提及率': f"{low_rate:.0f}%",
                '差距': f"{diff:+.0f}%"
            })
        c4_df = pd.DataFrame(c4_data)
        print(c4_df.to_string(index=False))
        print("\n  💡 业务含义：高付费玩家抱怨最多的问题，是商业化优化的首要目标")
    else:
        print("  ⚠️ 分组样本量不足（高付费或低付费 < 3），跳过 C4")
else:
    print("  ⚠️ 数据不足，跳过 C4")


# ---------- C5：玩家类型 × 联机意愿 ----------
print("\n" + "-" * 40)
print("【C5】玩家类型 × 联机意愿")
print("  业务问题：谁最期待联机？谁最反对联机？")
print("-" * 40)

if TYPE_COLS and WILLING_COL in df.columns:
    c5_data = []
    willing_high = ['非常想参与', '比较想参与']
    willing_low = ['不太想参与', '完全不想参与']

    for t in TYPE_COLS:
        type_n = len(df[df[t] == 1])
        if type_n >= 3:
            type_df = df[df[t] == 1]
            high_rate = len(type_df[type_df[WILLING_COL].isin(willing_high)]) / type_n * 100
            low_rate = len(type_df[type_df[WILLING_COL].isin(willing_low)]) / type_n * 100
            c5_data.append({
                '玩家类型': t.replace(' — 喜欢...', '').replace('党', '党')[:10],
                '人数': type_n,
                '想参与率': f"{high_rate:.0f}%",
                '不想参与率': f"{low_rate:.0f}%"
            })

    if c5_data:
        c5_df = pd.DataFrame(c5_data)
        print(c5_df.to_string(index=False))
        print("\n  💡 业务含义：想参与率高且不想参与率低的群体是联机玩法的核心用户")

        # 生成分组柱状图
        fig, ax = plt.subplots(figsize=(10, 6))
        x = range(len(c5_df))
        ax.bar(x, [float(r.replace('%', '')) for r in c5_df['想参与率']],
               width=0.35, label='想参与', color='#4ECDC4', align='center')
        ax.bar([i + 0.35 for i in x], [float(r.replace('%', '')) for r in c5_df['不想参与率']],
               width=0.35, label='不想参与', color='#FF6B6B', align='center')
        ax.set_xticks([i + 0.175 for i in x])
        ax.set_xticklabels(c5_df['玩家类型'])
        ax.set_ylabel('占比（%）')
        ax.set_title('各玩家类型联机意愿对比', fontsize=14)
        ax.legend()
        plt.tight_layout()
        plt.savefig('insight_type_willingness.png', dpi=300)
        print("  ✅ 已保存：insight_type_willingness.png")
    else:
        print("  ⚠️ 各类型样本量均 < 3，跳过 C5")
else:
    print("  ⚠️ 数据不足，跳过 C5")

print("\n" + "=" * 80)
print("✅ C1-C5 探索性分析完成！")
print("=" * 80)


# =============================================
# 第五部分：导出 Excel 汇总表
# =============================================
print("\n" + "=" * 80)
print("十五、导出表格到 Excel")
print("=" * 80)

with pd.ExcelWriter('分析结果汇总.xlsx', engine='openpyxl') as writer:

    # 5.1 玩家画像
    for col, sheet_name in [('1、您的性别是？', '性别'), ('2、您的年龄段是', '年龄'),
                            ('3、您目前的职业是', '职业'), ('4、您玩《明日方舟》多长时间了？', '游戏年龄'),
                            ('5、您一般每天花多少时间在《明日方舟》上？', '每日时长')]:
        if col in df.columns:
            series = df[col].value_counts()
            table = pd.DataFrame({'选项': series.index, '人数': series.values,
                                  '占比': [f"{v / total_n * 100:.1f}%" for v in series.values]})
            table.to_excel(writer, sheet_name=sheet_name, index=False)

    # 5.2 时间分配
    if time_data:
        pd.DataFrame(time_data).to_excel(writer, sheet_name='时间分配', index=False)

    # 5.3 玩家类型
    if type_data:
        pd.DataFrame(type_data).to_excel(writer, sheet_name='玩家类型', index=False)

    # 5.4 满意度
    if sat_data:
        pd.DataFrame(sat_data).to_excel(writer, sheet_name='满意度汇总', index=False)

    # 5.5 痛点
    if pain_data:
        pd.DataFrame(sorted(pain_data, key=lambda x: int(x['人数']), reverse=True)).to_excel(
            writer, sheet_name='痛点', index=False)

    # 5.6 付费
    if pay_data:
        pd.DataFrame(pay_data).to_excel(writer, sheet_name='付费分布', index=False)
    if motiv_data:
        pd.DataFrame(motiv_data).to_excel(writer, sheet_name='付费动机', index=False)

    # 5.7 NPS
    if nps_data:
        pd.DataFrame(nps_data).to_excel(writer, sheet_name='NPS', index=False)

    # 5.8 联机
    if will_data:
        pd.DataFrame(will_data).to_excel(writer, sheet_name='联机意愿', index=False)

    # 5.9 竞品数据
    if game_list:
        pd.DataFrame(game_data).to_excel(writer, sheet_name='竞品接触', index=False)

    # 5.10 其他内容列
    if other_cols:
        for col in other_cols:
            non_empty = df[df[col].notna() & (df[col] != '')]
            if len(non_empty) > 0:
                other_table = pd.DataFrame({
                    '用户编号': non_empty.index,
                    '来源列': col,
                    '内容': non_empty[col].values
                })
                sheet_name_clean = col.replace('_其他内容', '').replace('_', '')[:12]
                other_table.to_excel(writer, sheet_name=sheet_name_clean, index=False)

    # 5.11 H1 结果
    if h1_results:
        pd.DataFrame(h1_results).fillna(0).to_excel(writer, sheet_name='H1_类型×付费', index=False)

    # 5.12 H2 结果
    if h2_data:
        pd.DataFrame(h2_data).to_excel(writer, sheet_name='H2_痛点×NPS', index=False)

print("  ✅ 已导出：分析结果汇总.xlsx")


# =============================================
# 第六部分：生成可视化图表
# =============================================
print("\n" + "=" * 80)
print("十六、生成图表")
print("=" * 80)

# 6.1 满意度雷达图
if all(c in df.columns for c in SAT_COLS):
    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
    cats = ['美术', '剧情', '关卡', '音乐', '整体']
    vals = [df[c].mean() for c in SAT_COLS]
    angles = [n / len(cats) * 2 * pi for n in range(len(cats))]
    angles += angles[:1]
    vals_plot = vals + vals[:1]
    ax.plot(angles, vals_plot, linewidth=2, color='#FF6B6B')
    ax.fill(angles, vals_plot, alpha=0.25, color='#FF6B6B')
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(cats)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_title('满意度雷达图')
    plt.tight_layout()
    plt.savefig('satisfaction_radar.png', dpi=300)
    print("  ✅ satisfaction_radar.png")

# 6.2 痛点排名图
if pain_data:
    plt.figure(figsize=(10, 6))
    sorted_pain = sorted(pain_data, key=lambda x: int(x['人数']), reverse=True)
    names = [x['痛点'] for x in sorted_pain]
    values = [x['人数'] for x in sorted_pain]
    plt.barh(names, values, color='#4ECDC4')
    plt.xlabel('提及人数')
    plt.title('痛点排名')
    plt.tight_layout()
    plt.savefig('pain_rank.png', dpi=300)
    print("  ✅ pain_rank.png")

# 6.3 付费分布图
if PAY_COL in df.columns:
    plt.figure(figsize=(8, 8))
    pay_labels = list(df[PAY_COL].value_counts().index)
    pay_values = list(df[PAY_COL].value_counts().values)
    plt.pie(pay_values, labels=pay_labels, autopct='%1.1f%%', startangle=90,
            colors=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD'])
    plt.title('付费金额分布')
    plt.tight_layout()
    plt.savefig('pay_distribution.png', dpi=300)
    print("  ✅ pay_distribution.png")

# 6.4 竞品排名图
if '玩过游戏' in df.columns and game_list:
    plt.figure(figsize=(10, 6))
    sorted_games = sorted(game_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    games_sorted = [g for g, _ in sorted_games]
    counts_sorted = [c for _, c in sorted_games]
    plt.barh(games_sorted, counts_sorted, color='#4ECDC4')
    plt.xlabel('提及人数')
    plt.title('玩家提及游戏排名 TOP10', fontsize=14)
    plt.tight_layout()
    plt.savefig('competitor_rank.png', dpi=300)
    print("  ✅ competitor_rank.png")

print("\n" + "=" * 80)
print("✅ 全部完成！")
print("=" * 80)