import pandas as pd

def check_latest_bug_report():
    """检查修复后的Bug级别分析报告"""
    
    # 查找最新的报告文件
    import glob
    import os
    
    pattern = "output/Bug级别分析报告_*.xlsx"
    files = glob.glob(pattern)
    if not files:
        print("没有找到Bug级别分析报告文件")
        return
    
    # 获取最新的文件
    latest_file = max(files, key=os.path.getctime)
    print(f"检查文件: {latest_file}")
    
    try:
        # 读取Bug级别统计工作表
        df = pd.read_excel(latest_file, sheet_name="Bug级别统计")
        
        print("=== 修复后的Bug级别统计工作表 ===")
        print(f"数据形状: {df.shape}")
        print(f"列名: {list(df.columns)}")
        print("\n完整数据:")
        print(df.to_string())
        
        # 检查总计行
        total_row = df[df['文件名称'] == '总计']
        if not total_row.empty:
            print(f"\n=== 总计行验证 ===")
            total_data = total_row.iloc[0]
            print("总计行数据:")
            for col in df.columns:
                print(f"  {col}: {total_data[col]}")
            
            # 验证数据正确性
            print(f"\n=== 数据正确性验证 ===")
            expected_totals = {
                'A级': 32,
                'B级': 152, 
                'C级': 3,
                'S级': 0
            }
            
            for level, expected in expected_totals.items():
                if level in df.columns:
                    actual = total_data[level]
                    status = "✅ 正确" if actual == expected else "❌ 错误"
                    print(f"{level}: 期望={expected}, 实际={actual} {status}")
        
        # 检查是否还有Unnamed列
        unnamed_cols = [col for col in df.columns if 'Unnamed' in str(col)]
        if unnamed_cols:
            print(f"\n⚠️  仍然存在异常列: {unnamed_cols}")
        else:
            print(f"\n✅ 没有发现Unnamed异常列")
            
    except Exception as e:
        print(f"读取报告时出错: {e}")

if __name__ == "__main__":
    check_latest_bug_report()