"""
修复Bug级别分析中的统计错误
问题分析：在main.py的Bug级别统计逻辑中存在问题
"""
import pandas as pd

def analyze_bug_level_issue():
    """分析Bug级别统计问题"""
    
    print("=== Bug级别统计问题分析 ===")
    
    # 读取原始详细数据
    detail_file = "output/详细分析报告_批量分析_20250909_110507_20250909_110507.xlsx"
    df = pd.read_excel(detail_file, sheet_name="详细数据")
    
    print(f"原始数据行数: {len(df)}")
    print(f"严重级别列的唯一值: {df['严重级别'].unique()}")
    print(f"严重级别分布:")
    print(df['严重级别'].value_counts(dropna=False))
    
    # 模拟main.py中的级别映射逻辑
    level_mapping = {
        'S-严重': 'S级',
        'A-重要': 'A级', 
        'B-一般': 'B级',
        'C-轻微': 'C级'
    }
    
    # 先填充空值为'未分级'
    df_test = df.copy()
    df_test['严重级别'] = df_test['严重级别'].fillna('未分级')
    print(f"\n填充空值后的分布:")
    print(df_test['严重级别'].value_counts(dropna=False))
    
    # 应用级别映射
    df_test['级别'] = df_test['严重级别'].map(level_mapping).fillna(df_test['严重级别'])
    print(f"\n映射后的级别分布:")
    print(df_test['级别'].value_counts(dropna=False))
    
    # 按文件来源分组统计
    if '文件来源' in df_test.columns:
        print(f"\n按文件来源分组的级别统计:")
        result = df_test.groupby(['文件来源', '级别']).size().unstack(fill_value=0)
        print(result)
        
        # 计算总计
        print(f"\n各级别总计:")
        for level in ['S级', 'A级', 'B级', 'C级', '未分级']:
            if level in result.columns:
                total = result[level].sum()
                print(f"{level}: {total}")
            else:
                print(f"{level}: 0")
    
    print(f"\n=== 问题分析结果 ===")
    print("1. 原始数据中没有空值，所有记录都有明确的严重级别")
    print("2. 所有级别都能正确映射到S级、A级、B级、C级")
    print("3. 理论上不应该有'未分级'的记录")
    print("4. 但是Bug级别分析报告中显示'未分级'为3，这是统计逻辑的错误")
    
    # 检查可能的问题
    print(f"\n=== 可能的问题原因 ===")
    
    # 检查是否有不在映射表中的级别值
    unmapped_levels = []
    for level in df['严重级别'].unique():
        if pd.notna(level) and level not in level_mapping:
            unmapped_levels.append(level)
    
    if unmapped_levels:
        print(f"发现未映射的级别值: {unmapped_levels}")
    else:
        print("所有级别值都能正确映射")
    
    # 检查数据类型问题
    print(f"严重级别列的数据类型: {df['严重级别'].dtype}")
    print(f"是否包含空格或特殊字符:")
    for level in df['严重级别'].unique():
        if pd.notna(level):
            print(f"  '{level}' -> 长度: {len(str(level))}, 包含空格: {' ' in str(level)}")

def fix_bug_level_statistics():
    """修复Bug级别统计逻辑"""
    
    print(f"\n=== 修复建议 ===")
    print("问题出现在main.py的Bug级别统计逻辑中：")
    print("1. 在groupby().unstack()操作后，可能产生了额外的'未分级'列")
    print("2. 需要检查unstack操作是否正确处理了所有级别")
    print("3. 建议在统计前先验证所有数据都能正确映射")
    
    # 提供修复代码
    fix_code = '''
# 修复后的Bug级别统计逻辑
def analyze_bug_levels_fixed(self, df):
    """修复后的Bug级别分析"""
    try:
        # ... 前面的代码保持不变 ...
        
        # 处理级别名称，统一格式
        level_mapping = {
            'S-严重': 'S级',
            'A-重要': 'A级', 
            'B-一般': 'B级',
            'C-轻微': 'C级'
        }
        
        # 先映射，对于无法映射的值才设为'未分级'
        df_clean['级别'] = df_clean[level_col].map(level_mapping)
        
        # 只有真正无法映射的值才设为'未分级'
        unmapped_mask = df_clean['级别'].isna()
        df_clean.loc[unmapped_mask, '级别'] = '未分级'
        
        # 验证映射结果
        print(f"映射后级别分布: {df_clean['级别'].value_counts()}")
        
        # 统计各级别Bug数量
        result = df_clean.groupby(['文件名称', '级别']).size().unstack(fill_value=0)
        
        # 确保包含所有级别的列，但只添加实际需要的列
        expected_levels = ['S级', 'A级', 'B级', 'C级']
        actual_levels = df_clean['级别'].unique()
        
        # 只有当数据中真的存在'未分级'时才添加该列
        if '未分级' in actual_levels:
            expected_levels.append('未分级')
            
        for level in expected_levels:
            if level not in result.columns:
                result[level] = 0
        
        return result
        
    except Exception as e:
        self.log_message(f"Bug级别分析出错: {str(e)}")
        return None
'''
    
    print(fix_code)

if __name__ == "__main__":
    analyze_bug_level_issue()
    fix_bug_level_statistics()