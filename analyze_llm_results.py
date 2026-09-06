# =============================================
# 大模型文本分析结果 - 深度业务分析
# =============================================
# 【功能说明】
#   对「文本分析结果.xlsx」（大模型已标注情感/主题/建议价值）
#   进行业务层面的深度分析，输出优先级排序和业务洞察。
#
# 【分析流程】
#   1. 整体画像：情感分布、主题分布、建议价值分布
#   2. 交叉分析：情感 × 主题 交叉表 + 热力图
#   3. 优先级排序：
#      - P0（紧急问题）：情感=negative + 建议价值=high
#      - P1（待优化）：情感=negative + 建议价值=medium
#      - P1（高价值建议）：建议价值=high（不受情感限制）
#      - P2（一般反馈）：其余
#   4. 各主题的问题特征（负面占比、高价值占比）
#   5. 高价值建议原文汇总
#   6. 导出 Excel 汇总（完整标注、主题统计、P0问题、高价值建议）
#   7. 生成可视化图表（情感饼图、主题柱状图、优先级分布图）
#
# 【输入】  文本分析结果.xlsx（大模型输出）
# 【输出】  文本分析_深度汇总.xlsx + 3 张 PNG 图表
# =============================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
from collections import Counter
warnings.filterwarnings('ignore')

# ------------------------------
# 1. 全局设置
# ------------------------------
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

print("=" * 80)
print("大模型文本分析结果 - 深度业务分析")
print("=" * 80)

# ------------------------------
# 2. 读取数据
# ------------------------------
df = pd.read_excel('文本分析结果.xlsx')
total_n = len(df)
print(f"\n📌 共分析 {total_n} 条文本")
print(f"   来源列分布：{df['来源列'].nunique()} 种")
print(f"   涉及用户：{df['用户编号'].nunique()} 人")


# =============================================
# 第一部分：整体画像
# =============================================
print("\n" + "=" * 80)
print("一、整体画像")
print("=" * 80)

# 1.1 情感分布
print("\n【情感分布】")
sentiment_counts = df['情感'].value_counts()
for k, v in sentiment_counts.items():
    print(f"  {k}: {v}条 ({v/total_n*100:.1f}%)")

# 1.2 主题分布
print("\n【主题分布】")
theme_counts = df['主题'].value_counts()
for k, v in theme_counts.items():
    print(f"  {k}: {v}条 ({v/total_n*100:.1f}%)")

# 1.3 建议价值分布
print("\n【建议价值分布】")
value_counts = df['建议价值'].value_counts()
for k, v in value_counts.items():
    print(f"  {k}: {v}条 ({v/total_n*100:.1f}%)")


# =============================================
# 第二部分：交叉分析（情感 × 主题）
# =============================================
print("\n" + "=" * 80)
print("二、交叉分析（情感 × 主题）")
print("=" * 80)

# 2.1 交叉表
cross = pd.crosstab(df['情感'], df['主题'])
print("\n【情感 × 主题 交叉表】")
print(cross.to_string())

# 2.2 生成热力图
# 颜色映射：RdYlGn_r 使高值（红色区域）表示负面集中的主题
print("\n📊 生成热力图...")
plt.figure(figsize=(10, 6))
sns.heatmap(cross, annot=True, fmt='d', cmap='RdYlGn_r', center=5)
plt.title('情感 × 主题 热力图', fontsize=14)
plt.tight_layout()
plt.savefig('llm_sentiment_theme_heatmap.png', dpi=300)
print("  ✅ llm_sentiment_theme_heatmap.png")


# =============================================
# 第三部分：优先级排序（P0 / P1 / P2）
# =============================================
print("\n" + "=" * 80)
print("三、优先级排序（P0 / P1 / P2）")
print("=" * 80)

# 3.1 定义优先级规则
# P0：负面 + 高价值 → 最紧急，需立即响应
# P1-待优化：负面 + 中价值 → 需列入优化计划
# P1-高价值建议：正面/中性 + 高价值 → 值得采纳的建议
# P2：其他 → 一般反馈
def classify_priority(row):
    if row['情感'] == 'negative' and row['建议价值'] == 'high':
        return 'P0 - 紧急问题'
    elif row['情感'] == 'negative' and row['建议价值'] == 'medium':
        return 'P1 - 待优化'
    elif row['建议价值'] == 'high':
        return 'P1 - 高价值建议'
    else:
        return 'P2 - 一般反馈'

df['优先级'] = df.apply(classify_priority, axis=1)

print("\n【优先级分布】")
priority_counts = df['优先级'].value_counts()
for k, v in priority_counts.items():
    print(f"  {k}: {v}条 ({v/total_n*100:.1f}%)")

# 3.2 P0级问题详情
print("\n【🔴 P0级问题】（负面 + 高价值，最紧急）")
p0_df = df[df['优先级'] == 'P0 - 紧急问题']
if len(p0_df) > 0:
    for _, row in p0_df.iterrows():
        print(f"  - 用户{row['用户编号']} | {row['主题']} | {row['摘要']}")
        print(f"    原文：{row['原始文本'][:40]}...")
else:
    print("  ✅ 暂无P0级问题（或样本量不足）")

# 3.3 按主题统计 P0 问题分布
print("\n【P0问题按主题分布】")
p0_by_theme = p0_df['主题'].value_counts()
for k, v in p0_by_theme.items():
    print(f"  {k}: {v}条")


# =============================================
# 第四部分：各主题的问题特征
# =============================================
print("\n" + "=" * 80)
print("四、各主题的问题特征")
print("=" * 80)

# 计算每个主题的负面占比和高价值占比，识别问题集中领域
for theme in df['主题'].unique():
    theme_df = df[df['主题'] == theme]
    neg_ratio = len(theme_df[theme_df['情感'] == 'negative']) / len(theme_df) * 100
    high_ratio = len(theme_df[theme_df['建议价值'] == 'high']) / len(theme_df) * 100
    print(f"\n【{theme}】共 {len(theme_df)} 条")
    print(f"  负面占比：{neg_ratio:.1f}%")
    print(f"  高价值占比：{high_ratio:.1f}%")
    if len(theme_df) > 0:
        sample = theme_df.iloc[0]
        print(f"  典型反馈：{sample['摘要']}")


# =============================================
# 第五部分：高价值建议原文汇总
# =============================================
print("\n" + "=" * 80)
print("五、高价值建议原文汇总")
print("=" * 80)

high_df = df[df['建议价值'] == 'high']
print(f"\n共 {len(high_df)} 条高价值建议：")
for i, (_, row) in enumerate(high_df.iterrows(), 1):
    print(f"  {i}. [{row['主题']}] {row['原始文本']}")
    print(f"     → {row['摘要']}")


# =============================================
# 第六部分：导出 Excel 汇总
# =============================================
print("\n" + "=" * 80)
print("六、导出Excel汇总")
print("=" * 80)

with pd.ExcelWriter('文本分析_深度汇总.xlsx', engine='openpyxl') as writer:
    # 完整标注数据（含优先级）
    df.to_excel(writer, sheet_name='完整标注', index=False)
    
    # 各主题统计
    theme_summary = df.groupby('主题').agg({
        '情感': lambda x: x.count(),
        '原始文本': lambda x: len(x[x.str.len() > 0])
    }).rename(columns={'情感': '条数', '原始文本': '非空条数'})
    theme_summary.to_excel(writer, sheet_name='主题统计')
    
    # P0 紧急问题
    if len(p0_df) > 0:
        p0_df[['用户编号', '主题', '原始文本', '摘要']].to_excel(writer, sheet_name='P0紧急问题', index=False)
    
    # 高价值建议
    if len(high_df) > 0:
        high_df[['用户编号', '主题', '原始文本', '摘要']].to_excel(writer, sheet_name='高价值建议', index=False)

print("  ✅ 已导出：文本分析_深度汇总.xlsx")


# =============================================
# 第七部分：生成可视化图表
# =============================================
print("\n" + "=" * 80)
print("七、生成可视化图表")
print("=" * 80)

# 7.1 情感分布饼图
plt.figure(figsize=(8, 8))
colors = {'positive': '#4ECDC4', 'neutral': '#FFEAA7', 'negative': '#FF6B6B'}
plt.pie(sentiment_counts.values, labels=sentiment_counts.index, autopct='%1.1f%%',
        colors=[colors.get(k, '#95A5A6') for k in sentiment_counts.index])
plt.title('情感分布', fontsize=14)
plt.tight_layout()
plt.savefig('llm_sentiment_pie.png', dpi=300)
print("  ✅ llm_sentiment_pie.png")

# 7.2 主题分布柱状图（横向，便于阅读）
plt.figure(figsize=(10, 6))
theme_counts_sorted = theme_counts.sort_values(ascending=True)
plt.barh(theme_counts_sorted.index, theme_counts_sorted.values, color='#4ECDC4')
plt.xlabel('提及次数')
plt.title('主题分布', fontsize=14)
plt.tight_layout()
plt.savefig('llm_theme_bar.png', dpi=300)
print("  ✅ llm_theme_bar.png")

# 7.3 优先级分布柱状图
plt.figure(figsize=(10, 6))
priority_sorted = priority_counts.sort_values(ascending=False)
colors_priority = {'P0 - 紧急问题': '#FF6B6B', 'P1 - 待优化': '#FFEAA7', 
                   'P1 - 高价值建议': '#4ECDC4', 'P2 - 一般反馈': '#95A5A6'}
plt.bar(priority_sorted.index, priority_sorted.values,
        color=[colors_priority.get(k, '#95A5A6') for k in priority_sorted.index])
plt.xlabel('优先级')
plt.ylabel('条数')
plt.title('优先级分布', fontsize=14)
plt.xticks(rotation=15)
plt.tight_layout()
plt.savefig('llm_priority_bar.png', dpi=300)
print("  ✅ llm_priority_bar.png")

print("\n" + "=" * 80)
print("✅ 全部完成！")
print("=" * 80)