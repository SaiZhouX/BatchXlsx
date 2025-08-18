import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DataValidator:
    """数据验证类，用于检查数据完整性和一致性"""
    
    def __init__(self, output_folder="output"):
        """
        初始化数据验证器
        
        Args:
            output_folder (str): 输出文件夹路径
        """
        self.output_folder = Path(output_folder)
    
    def find_latest_report(self, prefix="详细分析报告"):
        """
        查找最新的报告文件
        
        Args:
            prefix (str): 文件名前缀
            
        Returns:
            Path: 最新报告文件的路径
        """
        try:
            # 查找所有匹配前缀的文件
            matching_files = list(self.output_folder.glob(f"{prefix}*.xlsx"))
            
            if not matching_files:
                logger.warning(f"未找到匹配 {prefix} 的文件")
                return None
            
            # 按修改时间排序，获取最新的文件
            latest_file = max(matching_files, key=os.path.getmtime)
            logger.info(f"找到最新报告文件: {latest_file}")
            
            return latest_file
            
        except Exception as e:
            logger.error(f"查找最新报告文件时出错: {str(e)}")
            return None
    
    def check_data_integrity(self, original_file, merged_file, filter_column, filter_value):
        """
        检查数据完整性
        
        Args:
            original_file (str): 原始文件路径
            merged_file (str): 合并文件路径
            filter_column (str): 过滤列名
            filter_value (str): 过滤值
            
        Returns:
            bool: 数据是否完整
        """
        try:
            print(f"检查数据完整性")
            print("="*60)
            
            # 1. 检查原始文件
            print(f"1. 原始文件检查: {original_file}")
            
            # 读取原始文件
            df_original = pd.read_excel(original_file)
            print(f"原始文件行数: {len(df_original)}")
            
            # 2. 检查合并文件
            print(f"\n2. 合并文件检查: {merged_file}")
            
            df_merged = pd.read_excel(merged_file)
            df_filtered = df_merged[df_merged[filter_column].str.contains(filter_value, na=False)]
            print(f"合并文件中过滤后行数: {len(df_filtered)}")
            
            # 3. 对比分析
            print(f"\n3. 对比分析:")
            print(f"原始文件: {len(df_original)}行")
            print(f"合并文件过滤后: {len(df_filtered)}行")
            print(f"差异: {len(df_original) - len(df_filtered)}行")
            
            if len(df_original) != len(df_filtered):
                print("❌ 数据有丢失!")
                return False
            else:
                print("✅ 数据完整!")
                return True
                
        except Exception as e:
            print(f"检查过程中出错: {str(e)}")
            return False
    
    def check_missing_data(self, merged_file, filter_column, filter_value):
        """
        检查丢失数据的原因
        
        Args:
            merged_file (str): 合并文件路径
            filter_column (str): 过滤列名
            filter_value (str): 过滤值
        """
        try:
            print("="*60)
            print(f"检查{filter_value}数据丢失原因")
            print("="*60)
            
            # 读取合并文件
            df = pd.read_excel(merged_file)
            
            # 筛选数据
            df_filtered = df[df[filter_column].str.contains(filter_value, na=False)]
            print(f"合并文件中{filter_value}数据总行数: {len(df_filtered)}")
            
            # 检查每一行的详细信息
            print(f"\n{filter_value}数据详细信息:")
            for i, row in df_filtered.iterrows():
                编号 = row.get('编号', 'N/A')
                级别 = row.get('严重级别', 'N/A')
                来源 = row.get('文件来源', 'N/A')
                bug类型 = row.get('bug类型', 'N/A')
                功能模块 = row.get('功能模块', 'N/A')
                print(f"  行{i}: 编号={编号}, 级别='{级别}', bug类型={bug类型}, 功能模块={功能模块}")
            
            # 统计级别分布
            print(f"\n{filter_value}数据级别分布:")
            level_column = next((col for col in df.columns if any(keyword in col.lower() for keyword in ['级别', 'level', '等级', 'priority', '严重', 'severity'])), None)
            
            if level_column:
                level_counts = df_filtered[level_column].value_counts(dropna=False)
                print(level_counts)
            
                # 检查是否有空值或特殊值
                print(f"\n检查空值情况:")
                for col in ['编号', level_column, 'bug类型', '功能模块']:
                    if col in df_filtered.columns:
                        null_count = df_filtered[col].isnull().sum()
                        print(f"  {col}: {null_count}个空值")
            else:
                print("未找到级别相关列")
            
        except Exception as e:
            print(f"检查过程中出错: {str(e)}")