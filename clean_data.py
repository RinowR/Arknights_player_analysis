# =============================================
# 《明日方舟》玩家问卷 - 数据清洗
# 功能：列名匹配、多选拆分、其他内容提取、游戏名称拆分、毒数据剔除
# =============================================

import pandas as pd
import numpy as np
import re
import warnings
warnings.filterwarnings('ignore')

print("=" * 60)
print("《明日方舟》问卷数据清洗")
print("=" * 60)

# =============================================
# 1. 读取数据
# =============================================
file_path = '《明日方舟》玩家游戏体验与满意度调查_28_28.xlsx'
df = pd.read_excel(file_path, header=0)
df.columns = df.columns.str.strip()

print(f"\n[1] 读取数据：{len(df)} 行，{len(df.columns)} 列")

# =============================================
# 2. 删除无关列（含邮箱列，保护隐私）
# =============================================
drop_cols = [
    '序号', '提交答卷时间', '所用时间', '来源', '来源详情', '来自IP', '总分',
    '23、如果愿意的话，请留下您的邮箱参与抽奖！'
]
drop_cols_existing = [c for c in drop_cols if c in df.columns]
df.drop(columns=drop_cols_existing, inplace=True)
print(f"[2] 删除无关列 {len(drop_cols_existing)} 列，剩余 {len(df.columns)} 列")

# =============================================
# 3. 重命名满意度列（第8题）
# =============================================
rename_map = {
    '8、请对以下《明日方舟》各方面进行满意度评价—角色立绘与美术风格': '满意_美术',
    '8、角色剧情与人设塑造': '满意_剧情',
    '8、关卡设计与策略性': '满意_关卡',
    '8、游戏音乐与音效': '满意_音乐',
    '8、游戏整体体验': '满意_整体'
}
rename_existing = {k: v for k, v in rename_map.items() if k in df.columns}
df.rename(columns=rename_existing, inplace=True)
print(f"[3] 重命名满意度列 {len(rename_existing)} 个")

# =============================================
# 4. 处理集成战略题（第11题）
# =============================================
def convert_rating(val):
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str):
        if '未参与' in val or '不评价' in val:
            return np.nan
        match = re.search(r'(\d+)', val)
        if match:
            return int(match.group(1))
    return np.nan

part_col = '11、您对“集成战略”（肉鸽）模式的参与度和满意度如何？—参与度'
sat_col = '11、满意度'
if part_col in df.columns:
    df['肉鸽_参与度'] = df[part_col].apply(convert_rating)
    df.drop(columns=[part_col], inplace=True)
if sat_col in df.columns:
    df['肉鸽_满意度'] = df[sat_col].apply(convert_rating)
    df.drop(columns=[sat_col], inplace=True)
print("[4] 处理集成战略题完成")

# =============================================
# 5. 拆分多选题（提取"其他内容"列 + 正常选项哑变量）
# =============================================
multi_choice_cols = [
    '6、您在《明日方舟》上主要把时间花在哪些方面？',
    '7、以下哪种描述最符合您的玩家类型？',
    '9、以下哪些是吸引您持续玩《明日方舟》的主要因素？',
    '10、您在游戏中遇到的最主要的挫折或不满是什么？',
    '14、您的付费主要用于哪些方面？',
    '15、促使您为《明日方舟》付费的主要因素有什么？',
    '19、除了《明日方舟》，您近期还经常玩以下哪些游戏？'
]

def split_multi_optimized(df, col_name, sep='┋', other_col_suffix='_其他内容'):
    """
    拆分多选题：
    1. 提取"其他〖xxx〗"中的内容 → {题号}_其他内容 列
    2. 其余选项 → 0/1 哑变量列
    3. 删除原始列
    """
    if col_name not in df.columns:
        return df
    
    # 提取题号作为前缀
    prefix = col_name.split('、')[0] if '、' in col_name else col_name[:2]
    
    # 提取"其他"内容的正则
    other_pattern = re.compile(r'其他[〖【（(]\s*(.*?)\s*[〗】）)]')
    
    def extract_other_text(val):
        if not isinstance(val, str):
            return ''
        if '其他' not in val:
            return ''
        match = other_pattern.search(val)
        if match:
            return match.group(1).strip()
        # 兼容其他格式：如 "其他：xxx"
        if '其他' in val:
            return val.replace('其他', '').strip('：:;；，,')
        return ''
    
    # 生成"其他内容"列
    other_col_name = prefix + other_col_suffix
    df[other_col_name] = df[col_name].apply(extract_other_text)
    
    # 获取所有非"其他"的选项（用于生成哑变量）
    all_options = set()
    for val in df[col_name].dropna():
        if isinstance(val, str):
            for opt in val.split(sep):
                opt = opt.strip()
                if not opt.startswith('其他'):
                    all_options.add(opt)
    
    # 为每个选项创建0/1哑变量
    for opt in all_options:
        df[opt] = df[col_name].str.contains(opt, na=False).astype(int)
    
    # 删除原始列
    df.drop(columns=[col_name], inplace=True)
    
    return df

print("\n开始拆分多选题...")
split_count = 0
for col in multi_choice_cols:
    before = df.shape[1]
    df = split_multi_optimized(df, col)
    if df.shape[1] > before:
        split_count += 1
        print(f"  ✅ {col[:20]}... → 已拆分")
print(f"[5] 拆分多选题 {split_count} 道")

# =============================================
# 6. 提取游戏名称（仅针对第19题"其他内容"）
#    生成"玩过游戏"列，以逗号分隔
# =============================================
print("\n" + "=" * 60)
print("🎮 提取「19_其他内容」中的游戏名称")
print("=" * 60)

def extract_game_names_from_19(df, other_col_name='19_其他内容'):
    """
    从19_其他内容列中提取游戏名称，合并为逗号分隔的文本
    生成一列：玩过游戏 = "原神, 崩坏星穹铁道, 边狱巴士"
    只针对第19题，其他题目的"其他内容"保留原样
    """
    # 游戏关键词映射
    game_keywords = {
        '原神': ['原神', 'genshin'],
        '崩坏星穹铁道': ['崩坏星穹铁道', '星穹铁道', '崩铁'],
        '鸣潮': ['鸣潮', 'wuthering'],
        '绝区零': ['绝区零', 'zzz'],
        '碧蓝航线': ['碧蓝航线', '碧蓝'],
        '重返未来1999': ['重返未来1999', '重返未来', '1999'],
        'FGO': ['fgo', 'Fate/Grand Order', '命运冠位指定'],
        '明日方舟终末地': ['终末地', 'zmd', '明日方舟终末地'],
        '边狱巴士': ['边狱巴士', 'limbus'],
        '少前2': ['少前2', '少女前线2', '追放'],
        '异环': ['异环'],
        '忘却前夜': ['忘却前夜'],
        '战争雷霆': ['战争雷霆', 'warthunder'],
        '英雄联盟': ['英雄联盟', 'lol'],
        '巫师三': ['巫师三', '巫师3', 'witcher'],
        '女神异闻录': ['p5r', '女神异闻录', 'persona'],
        '音乐游戏': ['音游', '音乐游戏', 'maimai', 'arcaea', 'phigros'],
    }
    
    if other_col_name not in df.columns:
        print(f"  ⚠️ 未找到列：{other_col_name}")
        return df
    
    def extract_games(val):
        """从文本中提取所有匹配的游戏名，返回逗号分隔字符串"""
        if not isinstance(val, str):
            return ''
        val_lower = val.lower()
        found_games = []
        for game_name, keywords in game_keywords.items():
            for kw in keywords:
                if kw.lower() in val_lower:
                    found_games.append(game_name)
                    break
        return ', '.join(found_games) if found_games else ''
    
    # 生成新列：玩过游戏
    df['玩过游戏'] = df[other_col_name].apply(extract_games)
    
    # 统计提取结果
    non_empty = df[df['玩过游戏'] != '']
    print(f"  ✅ 共 {len(non_empty)} 条记录提取到游戏名称")
    
    # 打印样本预览
    if len(non_empty) > 0:
        print("\n  预览（前5条）：")
        for idx, row in non_empty.head(5).iterrows():
            print(f"    [{idx}] {row[other_col_name]} → {row['玩过游戏']}")
    
    return df

df = extract_game_names_from_19(df, '19_其他内容')

# =============================================
# 7. 毒数据自动剔除（只剔除全1分）
# =============================================
print("\n" + "=" * 60)
print("🗑️ 毒数据自动检测与剔除")
print("=" * 60)

sat_cols = ['满意_美术', '满意_剧情', '满意_关卡', '满意_音乐', '满意_整体']
existing_sat = [c for c in sat_cols if c in df.columns]

toxic_count = 0
loyal_count = 0

if existing_sat:
    df['_std'] = df[existing_sat].std(axis=1, skipna=True)
    df['_mean'] = df[existing_sat].mean(axis=1, skipna=True)
    
    # 全1分 → 恶意差评 → 剔除
    toxic_condition = (df['_std'] == 0) & (df['_mean'] == 1.0)
    toxic_indices = df[toxic_condition].index.tolist()
    toxic_count = len(toxic_indices)
    
    # 全5分 → 忠实铁粉 → 保留（仅计数）
    loyal_condition = (df['_std'] == 0) & (df['_mean'] == 5.0)
    loyal_indices = df[loyal_condition].index.tolist()
    loyal_count = len(loyal_indices)
    
    if toxic_indices:
        df = df.drop(index=toxic_indices)
    
    df.drop(columns=['_std', '_mean'], inplace=True)

print(f"  全1分恶意样本 {toxic_count} 个（已剔除），全5分忠实样本 {loyal_count} 个（保留）")

# =============================================
# 8. 缺失值检查
# =============================================
missing = df.isnull().sum()
missing = missing[missing > 0]
if len(missing) > 0:
    print(f"[6] 发现 {len(missing)} 列存在缺失值（未处理）")
else:
    print("[6] 无缺失值")

# =============================================
# 9. 保存清洗后的数据
# =============================================
df.to_excel('清洗后问卷数据.xlsx', index=False)

print(f"\n[7] 清洗完成！")
print(f"  最终数据：{df.shape[0]} 行，{df.shape[1]} 列")
print(f"  保存位置：清洗后问卷数据.xlsx")

# =============================================
# 10. 输出各"其他内容"列的统计（便于检查）
# =============================================
print("\n" + "=" * 60)
print("📋 「其他内容」列统计")
print("=" * 60)

other_cols = [c for c in df.columns if c.endswith('_其他内容')]
for col in other_cols:
    non_empty = df[df[col].notna() & (df[col] != '')]
    print(f"  {col}: {len(non_empty)} 条非空")

if '玩过游戏' in df.columns:
    non_empty = df[df['玩过游戏'].notna() & (df['玩过游戏'] != '')]
    print(f"  玩过游戏: {len(non_empty)} 条非空（从19_其他内容提取）")

print("\n" + "=" * 60)
print("✅ 数据清洗完成！")
print("=" * 60)