import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime
from excel_processor import ExcelProcessor

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BugAnalyzer(ExcelProcessor):
    """Bug数据分析类，继承自ExcelProcessor"""
    
    def analyze_bug_levels(self, excel_file_path):
        """
        分析Excel文件中的Bug级别分布
        按照日期统计S级、A级、B级、C级Bug的数量
        
        Args:
            excel_file_path (str): Excel文件路径
            
        Returns:
            pd.DataFrame: 按日期和级别统计的结果
        """
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_file_path)
            
            print(f"正在分析文件: {excel_file_path}")
            print(f"数据总行数: {len(df)}")
            print("\n数据列名:")
            for i, col in enumerate(df.columns):
                print(f"{i+1}. {col}")
            
            # 显示前几行数据以便了解结构
            print("\n前5行数据预览:")
            print(df.head())
            
            # 检查是否有来源文件列和级别相关的列
            source_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['来源', 'source', '文件', 'file'])]
            level_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['级别', 'level', '等级', 'priority', '严重', 'severity'])]
            
            print(f"\n检测到的来源相关列: {source_columns}")
            print(f"检测到的级别相关列: {level_columns}")
            
            if not source_columns or not level_columns:
                print("\n请手动指定来源列和级别列:")
                print("可用的列:")
                for i, col in enumerate(df.columns):
                    print(f"{i}: {col}")
                return None
            
            # 使用第一个检测到的来源列和级别列
            source_col = source_columns[0]
            level_col = level_columns[0]
            
            print(f"\n使用来源列: {source_col}")
            print(f"使用级别列: {level_col}")
            
            # 数据预处理 - 只删除来源列为空的行，级别为空的给默认值
            print(f"\n数据预处理前行数: {len(df)}")
            df_clean = df.dropna(subset=[source_col])  # 只要求来源列不为空
            
            # 对级别为空的数据给默认值
            df_clean = df_clean.copy()  # 避免SettingWithCopyWarning
            df_clean[level_col] = df_clean[level_col].fillna('未分级')
            
            print(f"数据预处理后行数: {len(df_clean)}")
            
            # 统计级别分布
            level_distribution = df_clean[level_col].value_counts()
            print(f"\n级别分布统计:")
            print(level_distribution)
            
            # 从来源文件名中提取日期和测试人信息
            df_clean['日期'] = df_clean[source_col].apply(self.extract_date_and_tester_from_filename)
            
            # 过滤掉无法提取日期的记录
            df_clean = df_clean.dropna(subset=['日期'])
            
            print(f"\n提取到的日期: {sorted(df_clean['日期'].unique())}")
            
            # 处理级别名称，统一格式
            level_mapping = {
                'S-严重': 'S级',
                'A-重要': 'A级', 
                'B-一般': 'B级',
                'C-轻微': 'C级'
            }
            df_clean['级别'] = df_clean[level_col].map(level_mapping).fillna(df_clean[level_col])
            
            # 统计各级别Bug数量
            result = df_clean.groupby(['日期', '级别']).size().unstack(fill_value=0)
            
            # 确保包含S、A、B、C级别的列
            for level in ['S级', 'A级', 'B级', 'C级']:
                if level not in result.columns:
                    result[level] = 0
            
            # 重新排序列
            level_order = ['S级', 'A级', 'B级', 'C级']
            available_levels = [level for level in level_order if level in result.columns]
            other_levels = [col for col in result.columns if col not in level_order]
            
            result = result[available_levels + other_levels]
            
            # 显示结果
            print("\n" + "="*50)
            print("Bug 级别分布")
            print("="*50)
            
            # 创建格式化的表格
            print(f"{'日期':<10}", end="")
            for col in result.columns:
                print(f"{col:<8}", end="")
            print()
            print("-" * (10 + len(result.columns) * 8))
            
            for date, row in result.iterrows():
                print(f"{date:<10}", end="")
                for value in row:
                    print(f"{value:<8}", end="")
                print()
            
            # 保存结果到output文件夹
            output_file = os.path.join(self.output_folder, f"Bug级别分布统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            result.to_excel(output_file)
            print(f"\n统计结果已保存到: {output_file}")
            
            return result
            
        except Exception as e:
            print(f"分析过程中出现错误: {str(e)}")
            return None
    
    def extract_date_and_tester_from_filename(self, filename):
        """
        从文件名中提取日期和测试人信息
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的日期和测试人信息
        """
        date_str = self.extract_date_from_filename(filename)
        tester_name = self.extract_tester_from_filename(filename)
        
        # 如果有测试人姓名，则组合日期和姓名
        if date_str and tester_name:
            return f"{date_str}_{tester_name}"
        elif date_str:
            return date_str
        else:
            return None
    
    def check_tester_data(self, tester_name, original_file, merged_file):
        """
        检查特定测试人员的数据是否完整
        
        Args:
            tester_name (str): 测试人员姓名
            original_file (str): 原始文件路径
            merged_file (str): 合并文件路径
            
        Returns:
            bool: 数据是否完整
        """
        try:
            print(f"检查{tester_name}数据完整性")
            print("="*60)
            
            # 1. 检查原始文件
            print(f"1. 原始文件检查: {original_file}")
            
            # 读取问题记录工作表
            df_original = pd.read_excel(original_file, sheet_name='问题记录')
            print(f"原始文件行数: {len(df_original)}")
            
            # 统计级别
            level_counts = df_original['严重级别'].value_counts(dropna=False)
            print(f"原始文件级别分布: {level_counts}")
            print(f"原始文件总计: {level_counts.sum()}")
            
            # 2. 检查合并文件
            print(f"\n2. 合并文件检查: {merged_file}")
            
            df_merged = pd.read_excel(merged_file)
            df_tester = df_merged[df_merged['文件来源'].str.contains(tester_name, na=False)]
            print(f"合并文件中{tester_name}行数: {len(df_tester)}")
            
            # 统计级别
            merged_level_counts = df_tester['严重级别'].value_counts(dropna=False)
            print(f"合并文件级别分布: {merged_level_counts}")
            print(f"合并文件总计: {merged_level_counts.sum()}")
            
            # 3. 对比分析
            print(f"\n3. 对比分析:")
            print(f"原始文件: {level_counts.sum()}行")
            print(f"合并文件: {merged_level_counts.sum()}行")
            print(f"差异: {level_counts.sum() - merged_level_counts.sum()}行")
            
            if level_counts.sum() != merged_level_counts.sum():
                print("❌ 数据有丢失!")
                
                # 详细对比每个级别
                for level in level_counts.index:
                    original_count = level_counts.get(level, 0)
                    merged_count = merged_level_counts.get(level, 0)
                    if original_count != merged_count:
                        print(f"  {level}: 原始{original_count}个 -> 合并{merged_count}个 (差异: {original_count - merged_count})")
                
                return False
            else:
                print("✅ 数据完整!")
                return True
                
        except Exception as e:
            print(f"检查过程中出错: {str(e)}")
            return False