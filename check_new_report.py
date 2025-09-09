import pandas as pd

def check_bug_level_report():
    """检查最新的Bug级别分析报告"""
    
    file_path = "output/Bug级别分析报告_20250909_111846.xlsx"
    
    try:
        # 读取Bug级别统计工作表
        df = pd.read_excel(file_path, sheet_name="Bug级别统计")
        
        print("=== Bug级别统计工作表内容 ===")
        print(f"数据形状: {df.shape}")
        print(f"列名: {list(df.columns)}")
        print("\n完整数据:")
        print(df.to_string())
        
        # 检查总计行
        total_row = df[df['文件名称'] == '总计']
        if not total_row.empty:
            print(f"\n=== 总计行分析 ===")
            total_data = total_row.iloc[0]
            print("总计行数据:")
            for col in df.columns:
                print(f"  {col}: {total_data[col]}")
        
        # 检查各级别的实际统计
        print(f"\n=== 各文件级别统计验证 ===")
        level_cols = ['S级', 'A级', 'B级', 'C级', '未分级']
        for col in level_cols:
            if col in df.columns:
                # 排除总计行
                file_data = df[df['文件名称'] != '总计']
                total = file_data[col].sum()
                print(f"{col}: 各文件合计 = {total}")
        
    except Exception as e:
        print(f"读取报告时出错: {e}")

if __name__ == "__main__":
    check_bug_level_report()