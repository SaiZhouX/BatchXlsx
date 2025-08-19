"""
数据验证类，用于检查数据完整性和一致性
"""
import pandas as pd
import os
from pathlib import Path

from config_manager import config
from logger_config import LoggerConfig
from utils import FileUtils

class DataValidator:
    """数据验证类，用于检查数据完整性和一致性"""
    
    def __init__(self, output_folder=None):
        """
        初始化数据验证器
        
        Args:
            output_folder (str): 输出文件夹路径
        """
        self.output_folder = config.get_folder_path('output') if output_folder is None else Path(output_folder)
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
    
    def find_latest_report(self, prefix="详细分析报告"):
        """
        查找最新的报告文件
        
        Args:
            prefix (str): 文件名前缀
            
        Returns:
            Path: 最新报告文件的路径
        """
        pattern = f"{prefix}*.xlsx"
        latest_file = FileUtils.find_latest_file(str(self.output_folder), pattern)
        
        if latest_file:
            self.logger.info(f"找到最新报告文件: {latest_file.name}")
        else:
            self.logger.warning(f"未找到匹配 {prefix} 的文件")
        
        return latest_file
    
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
            self.logger.info("开始检查数据完整性")
            
            # 1. 检查原始文件
            self.logger.info(f"检查原始文件: {original_file}")
            df_original = pd.read_excel(original_file)
            self.logger.info(f"原始文件行数: {len(df_original)}")
            
            # 2. 检查合并文件
            self.logger.info(f"检查合并文件: {merged_file}")
            df_merged = pd.read_excel(merged_file)
            df_filtered = df_merged[df_merged[filter_column].str.contains(filter_value, na=False)]
            self.logger.info(f"合并文件中过滤后行数: {len(df_filtered)}")
            
            # 3. 对比分析
            difference = len(df_original) - len(df_filtered)
            self.logger.info(f"数据对比 - 原始: {len(df_original)}行, 合并过滤后: {len(df_filtered)}行, 差异: {difference}行")
            
            if difference != 0:
                self.logger.warning("数据有丢失!")
                return False
            else:
                self.logger.info("数据完整!")
                return True
                
        except Exception as e:
            self.logger.error(f"检查数据完整性时出错: {e}")
            return False
    
    def analyze_missing_data(self, merged_file, filter_column, filter_value):
        """
        分析丢失数据的原因
        
        Args:
            merged_file (str): 合并文件路径
            filter_column (str): 过滤列名
            filter_value (str): 过滤值
            
        Returns:
            dict: 分析结果
        """
        try:
            self.logger.info(f"开始分析{filter_value}数据丢失原因")
            
            # 读取合并文件
            df = pd.read_excel(merged_file)
            
            # 筛选数据
            df_filtered = df[df[filter_column].str.contains(filter_value, na=False)]
            self.logger.info(f"合并文件中{filter_value}数据总行数: {len(df_filtered)}")
            
            # 统计级别分布
            level_columns = config.get_level_columns()
            level_column = None
            
            for col in df.columns:
                if any(keyword in col.lower() for keyword in level_columns):
                    level_column = col
                    break
            
            analysis_result = {
                'total_rows': len(df_filtered),
                'level_distribution': {},
                'null_counts': {}
            }
            
            if level_column:
                level_counts = df_filtered[level_column].value_counts(dropna=False)
                analysis_result['level_distribution'] = level_counts.to_dict()
                self.logger.info(f"{filter_value}数据级别分布: {analysis_result['level_distribution']}")
            
                # 检查空值情况
                check_columns = ['编号', level_column, 'bug类型', '功能模块']
                for col in check_columns:
                    if col in df_filtered.columns:
                        null_count = df_filtered[col].isnull().sum()
                        analysis_result['null_counts'][col] = null_count
                        if null_count > 0:
                            self.logger.warning(f"{col}列有{null_count}个空值")
            else:
                self.logger.warning("未找到级别相关列")
            
            return analysis_result
            
        except Exception as e:
            self.logger.error(f"分析数据时出错: {e}")
            return None
    
    def validate_report_structure(self, report_file):
        """
        验证报告文件结构
        
        Args:
            report_file (str): 报告文件路径
            
        Returns:
            dict: 验证结果
        """
        try:
            self.logger.info(f"验证报告结构: {Path(report_file).name}")
            
            # 读取所有工作表
            excel_file = pd.ExcelFile(report_file)
            sheets = excel_file.sheet_names
            
            result = {
                'valid': True,
                'sheet_count': len(sheets),
                'sheet_names': sheets,
                'issues': []
            }
            
            # 检查必需的工作表
            required_sheets = ['详细数据', '分析统计']
            for sheet in required_sheets:
                if sheet not in sheets:
                    result['valid'] = False
                    result['issues'].append(f"缺少必需工作表: {sheet}")
            
            # 检查每个工作表的数据
            for sheet_name in sheets:
                try:
                    df = pd.read_excel(report_file, sheet_name=sheet_name)
                    if df.empty:
                        result['issues'].append(f"工作表 '{sheet_name}' 为空")
                    else:
                        self.logger.info(f"工作表 '{sheet_name}': {len(df)} 行 {len(df.columns)} 列")
                except Exception as e:
                    result['valid'] = False
                    result['issues'].append(f"读取工作表 '{sheet_name}' 失败: {e}")
            
            if result['valid']:
                self.logger.info("报告结构验证通过")
            else:
                self.logger.warning(f"报告结构验证失败: {result['issues']}")
            
            return result
            
        except Exception as e:
            self.logger.error(f"验证报告结构时出错: {e}")
            return {'valid': False, 'issues': [str(e)]}