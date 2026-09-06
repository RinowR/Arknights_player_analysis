# =============================================
# 《明日方舟》玩家问卷 - 大模型文本分析（完整修正版）
# 功能：一次性调用API分析所有文本，每条反馈独立输出
# 策略：激活内部知识 + 语境增强 + 后处理兜底
# 更新：排除19_其他内容（游戏名称），避免无效分析
# =============================================

import pandas as pd
import os
import json
import re
import time
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

# =============================================
# 1. 配置客户端
# =============================================

API_KEY = os.getenv('LLM_API_KEY')
BASE_URL = os.getenv('LLM_BASE_URL', 'https://www.cctq.ai/v1')
MODEL = os.getenv('LLM_MODEL', 'gpt-5.6-luna')

if not API_KEY:
    print("❌ 错误：未找到 LLM_API_KEY，请在 .env 文件中配置")
    exit(1)

client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

print("=" * 60)
print("《明日方舟》问卷 - 大模型文本分析（完整修正版）")
print("=" * 60)
print(f"使用模型：{MODEL}")
print(f"API地址：{BASE_URL}")
print("=" * 60)


# =============================================
# 2. 读取数据
# =============================================
df = pd.read_excel('清洗后问卷数据.xlsx')
print(f"\n📌 读取数据：{len(df)} 份问卷")


# =============================================
# 3. 提取需要分析的文本（⭐ 排除 19_其他内容）
# =============================================

# 3.1 提取所有"其他内容"列，但排除 19_其他内容（游戏名称）
other_cols = [c for c in df.columns if c.endswith('_其他内容') and c != '19_其他内容']

# 3.2 提取Q17（情感词）和Q22（改进建议）
text_cols = [
    '17、请用一个词形容您对《明日方舟》的整体情感：',
    '22、您对《明日方舟》有什么改进建议或想对鹰角网络说的话？'
]

# 3.3 汇总所有需要分析的文本
analysis_items = []

# 处理其他内容列（已排除 19_其他内容）
for col in other_cols:
    for idx, val in df[col].items():
        if pd.notna(val) and str(val).strip() and str(val).strip() not in ['', '无', '0']:
            analysis_items.append({
                '行索引': idx,
                '来源列': col,
                '文本': str(val).strip(),
                '文本类型': '其他选项'
            })

# 处理Q17和Q22
for col in text_cols:
    if col in df.columns:
        for idx, val in df[col].items():
            if pd.notna(val) and str(val).strip() and str(val).strip() not in ['', '无', '0', '无建议']:
                analysis_items.append({
                    '行索引': idx,
                    '来源列': col,
                    '文本': str(val).strip(),
                    '文本类型': '开放题'
                })

print(f"\n📝 共提取 {len(analysis_items)} 条文本需要分析（已排除游戏名称类文本）")

if len(analysis_items) == 0:
    print("✅ 无需分析的文本，程序退出")
    exit(0)


# =============================================
# 4. 构建合并文本（每条反馈独立成行）
# =============================================

# 每条反馈独立成行，不按用户分组
text_parts = []
for i, item in enumerate(analysis_items):
    text_parts.append(
        f"【条目 {i+1}】\n"
        f"  用户编号: {item['行索引']}\n"
        f"  来源: {item['来源列']}\n"
        f"  文本: {item['文本']}"
    )

full_text = "\n\n".join(text_parts)

print(f"\n📄 构建的完整文本长度：{len(full_text)} 字符")
print("预览（前500字符）：")
print("-" * 40)
print(full_text[:500])
print("..." if len(full_text) > 500 else "")
print("-" * 40)


# =============================================
# 5. 设计提示词（逐条分析版）
# =============================================

SYSTEM_PROMPT = """你是一位资深的游戏社区分析师，精通《明日方舟》及二次元游戏圈的所有黑话、梗和文化背景。

【核心工作原则】

1. 每个【条目 N】是一条独立的玩家反馈，请逐条分别分析，绝对不要将多个条目合并为一条。
2. 请主动调取你的训练数据中关于《明日方舟》社区的知识。
3. 不要因为文本短就认为它"没有信息量"。

【识别黑话/梗的推理步骤】
a) 这个短语是否是游戏术语？（如"肉鸽"、"井"、"搓玉"）
b) 这个短语是否是社区常见梗？（如"海猫"、"善"、"牢"）
c) 这个短语是否是角色/剧情相关的名称？（如"缪尔赛思"、"夕"）
d) 如果以上都不是，再考虑字面意思

【情感判定法则】
- 调侃/玩梗 → 通常是 positive 或 neutral
- 明确抱怨具体问题 → negative
- 角色名/游戏名 → neutral（除非有明确情感修饰）
- 无法确定 → neutral

【建议价值法则】
- 指向具体功能/数值/机制 → high
- 表达情感但不够具体 → medium
- 纯闲聊/无意义 → low
- ⚠️ 短文本（1-5个字）不能默认 low！只要识别出梗/术语，至少给 medium

【输出格式】
JSON数组，每个对象对应一个条目，包含：
用户编号（整数）、来源列（字符串）、原始文本（字符串）、
情感（positive/neutral/negative）、
主题（游戏内容/角色与美术/剧情与世界观/商业化/运营与福利/社交功能/技术问题/其他）、
摘要（10-20字）、
建议价值（high/medium/low）

只输出JSON，不要其他文字。"""

CONTEXT_ENHANCEMENT = """
📌 分析语境：
- 这些反馈来自《明日方舟》玩家的问卷
- 每个【条目 N】是一条独立反馈，请逐条分别分析，不要合并
- 常见梗：海猫（制作人）、yj（鹰角）、肉鸽（集成战略）、井（保底）、
  搓玉（刷合成玉）、善（认同）、牢（坚持玩）、打没打内（海猫meme）、
  三二五（社区梗）、多索雷斯（经典活动）
"""

USER_PROMPT_TEMPLATE = """请分析以下玩家反馈列表。每个条目用【条目 X】标记，每个条目是一条独立的反馈，请逐条分别分析，不要合并。

【分析语境】
{CONTEXT_ENHANCEMENT}

【重要】输出必须是一个JSON数组，数组长度必须等于条目数量。只输出JSON，不要其他文字。

反馈列表：
{full_text}"""


# =============================================
# 6. 调用API（只调用1次）
# =============================================

print("\n" + "=" * 60)
print("🚀 正在调用API分析所有文本（仅需1次调用）...")
print("=" * 60)

try:
    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": USER_PROMPT_TEMPLATE.format(
                CONTEXT_ENHANCEMENT=CONTEXT_ENHANCEMENT,
                full_text=full_text
            )}
        ],
        temperature=0.3,
        max_tokens=4096
    )
    
    result_text = response.choices[0].message.content.strip()
    print(f"\n✅ API调用完成，响应长度：{len(result_text)} 字符")

except Exception as e:
    print(f"❌ API调用失败：{e}")
    exit(1)


# =============================================
# 7. 解析JSON结果
# =============================================

print("\n" + "=" * 60)
print("📊 解析分析结果...")
print("=" * 60)

result_text = re.sub(r'```json\s*', '', result_text)
result_text = re.sub(r'```\s*', '', result_text)

try:
    parsed = json.loads(result_text)
    
    if not isinstance(parsed, list):
        print(f"⚠️ 返回的结果不是数组，而是 {type(parsed)}")
        print("尝试提取数组部分...")
        match = re.search(r'\[.*\]', result_text, re.DOTALL)
        if match:
            parsed = json.loads(match.group(0))
        else:
            raise ValueError("无法提取JSON数组")
    
    print(f"✅ 成功解析 {len(parsed)} 条分析结果")

except json.JSONDecodeError as e:
    print(f"❌ JSON解析失败：{e}")
    print(f"原始响应内容（前500字符）：\n{result_text[:500]}...")
    exit(1)


# =============================================
# 8. 构建结果DataFrame（修正重复行和合并问题）
# =============================================

results = []
for i, item in enumerate(parsed):
    results.append({
        '用户编号': item.get('用户编号', '未知'),
        '来源列': item.get('来源列', '未知'),
        '原始文本': item.get('原始文本', ''),
        '情感': item.get('情感', '未知'),
        '主题': item.get('主题', '其他'),
        '摘要': item.get('摘要', ''),
        '建议价值': item.get('建议价值', 'low'),
        '序号': i + 1
    })

result_df = pd.DataFrame(results)

# 用序号一一对应合并文本类型和行索引，避免笛卡尔积
orig_df = pd.DataFrame(analysis_items).reset_index(drop=True)
result_df = result_df.sort_values('序号').reset_index(drop=True)
result_df['行索引'] = orig_df['行索引']
result_df['文本类型'] = orig_df['文本类型']

# 调整列顺序
result_df = result_df[['用户编号', '行索引', '来源列', '文本类型', '原始文本', '情感', '主题', '摘要', '建议价值', '序号']]


# =============================================
# 8.5 后处理：通用修正（不硬编码特定词条）
# =============================================

def post_process_results(df):
    """
    通用后处理修正，不依赖特定词条列表：
    1. 短文本（<=5字符）且建议价值为low → 提升为medium
    2. 短文本且情感为neutral → 如果含中文且不含明显负面词，倾向positive
    3. 主题为"其他"但文本含游戏实体词 → 重归类为"游戏内容"
    """
    negative_words = ['差', '烂', '黑', '死', '骂', '批评', '不满', '失望', '垃圾', '恶心']
    game_entities = ['海猫', 'yj', '鹰角', '肉鸽', '多索雷斯', '缪尔赛思', '夕', '模组', 
                     '终末地', 'zmd', '集成战略', '井', '保底', '搓玉', '诗怀雅']
    
    for idx, row in df.iterrows():
        text = str(row['原始文本']).strip()
        
        # 规则1：短文本且被判定为low → 提升为medium
        if len(text) <= 5 and row['建议价值'] == 'low':
            if any('\u4e00' <= c <= '\u9fff' for c in text):
                df.at[idx, '建议价值'] = 'medium'
                if row['情感'] == 'neutral':
                    if not any(w in text for w in negative_words):
                        df.at[idx, '情感'] = 'positive'
        
        # 规则2：文本包含游戏实体词且主题为"其他" → 重归类
        if row['主题'] == '其他' and any(e in text for e in game_entities):
            df.at[idx, '主题'] = '游戏内容'
            if row['建议价值'] == 'low':
                df.at[idx, '建议价值'] = 'medium'
    
    return df

result_df = post_process_results(result_df)

# 保存
result_df.to_excel('文本分析结果.xlsx', index=False)

print(f"\n📄 已保存：文本分析结果.xlsx（共 {len(result_df)} 条，无重复行，每条独立）")


# =============================================
# 9. 生成分析报告
# =============================================
print("\n" + "=" * 60)
print("📊 分析报告")
print("=" * 60)

print("\n【情感分布】")
sentiment_counts = result_df['情感'].value_counts()
for k, v in sentiment_counts.items():
    print(f"  {k}: {v}条 ({v/len(result_df)*100:.1f}%)")

print("\n【主题分布】")
theme_counts = result_df['主题'].value_counts()
for k, v in theme_counts.items():
    print(f"  {k}: {v}条 ({v/len(result_df)*100:.1f}%)")

print("\n【建议价值分布】")
value_counts = result_df['建议价值'].value_counts()
for k, v in value_counts.items():
    print(f"  {k}: {v}条 ({v/len(result_df)*100:.1f}%)")

print("\n【情感 × 主题交叉表】")
cross = pd.crosstab(result_df['情感'], result_df['主题'])
print(cross.to_string())

print("\n【高价值建议（建议价值=high）】")
high_value = result_df[result_df['建议价值'] == 'high']
if len(high_value) > 0:
    for _, row in high_value.iterrows():
        print(f"  - 用户{row['用户编号']} | {row['摘要']}")
        print(f"    (来源: {row['原始文本'][:40]}...)")
else:
    print("  暂无高价值建议")

print("\n【负面情感 + 高价值建议（P0级问题）】")
p0 = result_df[(result_df['情感'] == 'negative') & (result_df['建议价值'] == 'high')]
if len(p0) > 0:
    for _, row in p0.iterrows():
        print(f"  - 用户{row['用户编号']} | {row['主题']} | {row['摘要']}")
else:
    print("  暂无")

print("\n" + "=" * 60)
print("✅ 全部完成！")
print("=" * 60)