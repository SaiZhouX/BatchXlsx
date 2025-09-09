import pandas as pd
import numpy as np

def analyze_bug_level_report():
    """分析Bug级别报告中的数据不一致问题"""
    
    file_path = "output/Bug级别分析报告_20250909_110516.xlsx"
    
    try:
        # 读取Excel文件的所有工作表
        excel_file = pd.ExcelFile(file_path)
        print(f"工作表列表: {excel_file.sheet_names}")
        
        # 读取每个工作表
        for sheet_name in excel_file.sheet_names:
            print(f"\n{'='*50}")
            print(f"工作表: {sheet_name}")
            print(f"{'='*50}")
            
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"数据形状: {df.shape}")
            print(f"列名: {list(df.columns)}")
            
            # 显示前几行数据
            print("\n前10行数据:")
            print(df.head(10))
            
            # 如果是统计相关的工作表，进行详细分析
            if '统计' in sheet_name or '分析' in sheet_name:
                print(f"\n详细数据分析:")
                print(df.to_string())
                
                # 检查是否有数值列
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    print(f"\n数值列统计:")
                    for col in numeric_cols:
                        print(f"{col}: 总和={df[col].sum()}, 非零值={df[col][df[col] != 0].count()}")
            
            # 如果包含严重级别相关数据
            if any('严重' in str(col) or '级别' in str(col) for col in df.columns):
                print(f"\n严重级别相关分析:")
                for col in df.columns:
                    if '严重' in str(col) or '级别' in str(col):
                        print(f"列 '{col}' 的唯一值:")
                        print(df[col].value_counts(dropna=False))
                        
        # 特别检查原始数据中的严重级别分布
        print(f"\n{'='*60}")
        print("检查原始数据中的严重级别分布")
        print(f"{'='*60}")
        
        # 尝试读取详细数据工作表
        try:
            detail_df = pd.read_excel("output/详细分析报告_批量分析_20250909_110507_20250909_110507.xlsx", sheet_name="详细数据")
            print(f"详细数据形状: {detail_df.shape}")
            
            # 查找严重级别列
            severity_cols = [col for col in detail_df.columns if '严重' in str(col) or '级别' in str(col)]
            print(f"严重级别相关列: {severity_cols}")
            
            for col in severity_cols:
                print(f"\n列 '{col}' 的分布:")
                value_counts = detail_df[col].value_counts(dropna=False)
                print(value_counts)
                
                # 检查空值和异常值
                null_count = detail_df[col].isnull().sum()
                empty_count = (detail_df[col] == '').sum()
                print(f"空值数量: {null_count}")
                print(f"空字符串数量: {empty_count}")
                
                # 显示一些具体的值
                print(f"前20个值的样本:")
                print(detail_df[col].head(20).tolist())
                
        except Exception as e:
            print(f"读取详细数据时出错: {e}")
            
    except Exception as e:
        print(f"分析报告时出错: {e}")

if __name__ == "__main__":
    analyze_bug_level_report()