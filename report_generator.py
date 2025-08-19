"""
报告生成器基类，提供统一的报告生成功能
"""
import pandas as pd
from pathlib import Path
from datetime import datetime

from config_manager import config
from logger_config import LoggerConfig
from utils import DataUtils, ExcelUtils

class ReportGenerator:
    """报告生成器基类"""
    
    def __init__(self):
        self.output_folder = config.get_folder_path('output')
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
        
        # 确保输出文件夹存在
        self.output_folder.mkdir(exist_ok=True)
    
    def clean_data(self, df):
        """
        清理数据
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        if df.empty:
            return df
        
        try:
            original_shape = df.shape
            
            # 使用工具函数清理数据
            df = DataUtils.remove_useless_columns(df)
            df = DataUtils.clean_empty_rows(df)
            
            cleaned_shape = df.shape
            self.logger.info(f"数据清理完成: 原有 {original_shape[0]} 行 {original_shape[1]} 列，"
                           f"清理后 {cleaned_shape[0]} 行 {cleaned_shape[1]} 列")
            
            return df
            
        except Exception as e:
            self.logger.error(f"数据清理时出错: {e}")
            return df
    
    def generate_statistics(self, df, source_info=None):
        """
        生成统计数据
        
        Args:
            df (pd.DataFrame): 数据
            source_info (str): 数据源信息
            
        Returns:
            list: 统计数据列表
        """
        try:
            stats_data = []
            
            # 基本信息
            stats_data.append(['处理时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            if source_info:
                stats_data.append(['数据源', source_info])
            
            # 数据规模
            stats_data.append(['', ''])  # 空行分隔
            stats_data.append(['=== 数据规模 ===', ''])
            stats_data.append(['总行数', len(df)])
            stats_data.append(['总列数', len(df.columns)])
            
            if not df.empty:
                # 数据质量
                stats_data.append(['', ''])
                stats_data.append(['=== 数据质量 ===', ''])
                
                # 缺失值统计
                missing_counts = df.isnull().sum()
                total_missing = missing_counts.sum()
                stats_data.append(['总缺失值', total_missing])
                
                if total_missing > 0:
                    missing_rate = (total_missing / (len(df) * len(df.columns))) * 100
                    stats_data.append(['缺失值比例', f'{missing_rate:.2f}%'])
                    
                    # 各列缺失值详情
                    stats_data.append(['', ''])
                    stats_data.append(['各列缺失值详情', ''])
                    for col, missing_count in missing_counts.items():
                        if missing_count > 0:
                            col_missing_rate = (missing_count / len(df)) * 100
                            stats_data.append([f'  {col}', f'{missing_count} ({col_missing_rate:.1f}%)'])
                
                # 重复值统计
                duplicate_count = df.duplicated().sum()
                stats_data.append(['重复行数', duplicate_count])
                if duplicate_count > 0:
                    duplicate_rate = (duplicate_count / len(df)) * 100
                    stats_data.append(['重复行比例', f'{duplicate_rate:.2f}%'])
                
                # 数据类型统计
                stats_data.append(['', ''])
                stats_data.append(['=== 数据类型 ===', ''])
                type_counts = df.dtypes.value_counts()
                for dtype, count in type_counts.items():
                    stats_data.append([f'{dtype} 类型列数', count])
                
                # 业务统计（如果包含特定列）
                business_stats = self._generate_business_statistics(df)
                if business_stats:
                    stats_data.extend(business_stats)
            
            return stats_data
            
        except Exception as e:
            self.logger.error(f"生成统计数据时出错: {e}")
            return [['错误', str(e)]]
    
    def _generate_business_statistics(self, df):
        """
        生成业务相关统计
        
        Args:
            df (pd.DataFrame): 数据
            
        Returns:
            list: 业务统计数据
        """
        business_stats = []
        
        try:
            # Bug相关统计
            if '严重级别' in df.columns:
                business_stats.append(['', ''])
                business_stats.append(['=== Bug级别统计 ===', ''])
                
                severity_counts = df['严重级别'].value_counts()
                for severity, count in severity_counts.items():
                    percentage = (count / len(df)) * 100
                    business_stats.append([severity, f'{count} ({percentage:.1f}%)'])
            
            if 'bug类型' in df.columns:
                business_stats.append(['', ''])
                business_stats.append(['=== Bug类型统计 ===', ''])
                
                type_counts = df['bug类型'].value_counts()
                for bug_type, count in type_counts.items():
                    percentage = (count / len(df)) * 100
                    business_stats.append([bug_type, f'{count} ({percentage:.1f}%)'])
            
            if '修复状态' in df.columns:
                business_stats.append(['', ''])
                business_stats.append(['=== 修复状态统计 ===', ''])
                
                status_counts = df['修复状态'].value_counts()
                for status, count in status_counts.items():
                    percentage = (count / len(df)) * 100
                    business_stats.append([status, f'{count} ({percentage:.1f}%)'])
            
            if '功能模块' in df.columns:
                business_stats.append(['', ''])
                business_stats.append(['=== 功能模块统计 ===', ''])
                
                module_counts = df['功能模块'].value_counts()
                # 只显示前10个最多的模块
                top_modules = module_counts.head(10)
                for module, count in top_modules.items():
                    percentage = (count / len(df)) * 100
                    business_stats.append([module, f'{count} ({percentage:.1f}%)'])
                
                if len(module_counts) > 10:
                    other_count = module_counts.tail(len(module_counts) - 10).sum()
                    other_percentage = (other_count / len(df)) * 100
                    business_stats.append(['其他模块', f'{other_count} ({other_percentage:.1f}%)'])
            
        except Exception as e:
            self.logger.error(f"生成业务统计时出错: {e}")
        
        return business_stats
    
    def generate_unified_report(self, df, report_name_prefix):
        """
        生成统一格式的Excel报告（2个标签页）
        
        Args:
            df (pd.DataFrame): 数据
            report_name_prefix (str): 报告名称前缀
            
        Returns:
            Path: 生成的报告文件路径
        """
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f"详细分析报告_{report_name_prefix}_{timestamp}.xlsx"
            report_path = self.output_folder / report_filename
            
            sheets_data = {}
            
            # 第一个标签页：详细数据
            if not df.empty and len(df.columns) > 0:
                sheets_data['详细数据'] = df
                self.logger.info(f"详细数据工作表: {len(df)} 行 {len(df.columns)} 列")
            else:
                # 如果数据为空，创建说明DataFrame
                empty_data = pd.DataFrame({
                    '数据说明': [
                        '数据处理结果',
                        f'处理时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
                        '数据状态: 处理后数据为空',
                        '可能原因: 原始文件无有效数据或数据被过滤',
                        '建议: 请检查原始文件是否包含有效数据'
                    ]
                })
                sheets_data['详细数据'] = empty_data
                self.logger.warning("数据为空，创建说明工作表")
            
            # 第二个标签页：分析统计
            stats_data = self.generate_statistics(df, report_name_prefix)
            stats_df = pd.DataFrame(stats_data, columns=['统计项', '值'])
            sheets_data['分析统计'] = stats_df
            
            # 保存Excel文件
            if ExcelUtils.save_excel_with_sheets(str(report_path), sheets_data):
                self.logger.info(f"已生成统一报告: {report_filename}")
                return report_path
            else:
                self.logger.error(f"保存统一报告失败: {report_filename}")
                return None
                
        except Exception as e:
            self.logger.error(f"生成统一报告时出错: {e}")
            return None