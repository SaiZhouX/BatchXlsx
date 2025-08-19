"""
报告生成器 - 提取单个分析和批量分析的共同功能
"""
import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime

class ReportGenerator:
    """报告生成器基类，包含单个分析和批量分析的共同功能"""
    
    def __init__(self, output_folder="output"):
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)
        self.logger = logging.getLogger(__name__)
    
    def clean_data(self, df):
        """
        清理数据，删除无用列和空行
        
        Args:
            df (pd.DataFrame): 输入的DataFrame
            
        Returns:
            pd.DataFrame: 清理后的DataFrame
        """
        if df.empty:
            return df
        
        # 删除类似"Unnamed: X"的列
        unnamed_cols = [col for col in df.columns if isinstance(col, str) and col.startswith('Unnamed:')]
        if unnamed_cols:
            df = df.drop(columns=unnamed_cols)
            self.logger.info(f"删除了无用列: {unnamed_cols}")
        
        # 删除完全为空的行和列
        df_cleaned = df.dropna(how='all')  # 只删除完全空的行
        df_cleaned = df_cleaned.dropna(axis=1, how='all')  # 只删除完全空的列
        
        # 如果清理后数据为空，返回原始数据（去除Unnamed列）
        if df_cleaned.empty:
            df_cleaned = df.drop(columns=unnamed_cols) if unnamed_cols else df
        
        return df_cleaned
    
    def generate_unified_report(self, df, report_name, source_info=None, original_df=None, original_rows=None, original_cols=None):
        """
        生成统一格式的Excel分析报告（2个标签页）
        
        Args:
            df (pd.DataFrame): 清理后的数据
            report_name (str): 报告名称
            source_info (str): 数据源信息
            original_df (pd.DataFrame): 原始数据（可选）
            original_rows (int): 原始行数（可选）
            original_cols (int): 原始列数（可选）
            
        Returns:
            str: 生成的报告文件路径
        """
        try:
            # 生成报告文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f"{report_name}_{timestamp}.xlsx"
            report_path = self.output_folder / report_filename
            
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                # 第一个工作表：详细数据
                self._write_detailed_data_sheet(writer, df, original_df, original_rows, original_cols)
                
                # 第二个工作表：分析统计
                self._write_analysis_stats_sheet(writer, df, source_info, original_df, original_rows, original_cols)
            
            self.logger.info(f"已生成统一报告: {os.path.basename(report_path)}")
            return str(report_path)
            
        except Exception as e:
            self.logger.error(f"生成统一报告失败: {str(e)}")
            return None
    
    def _write_detailed_data_sheet(self, writer, df, original_df=None, original_rows=None, original_cols=None):
        """写入详细数据工作表"""
        if not df.empty and len(df.columns) > 0:
            df.to_excel(writer, sheet_name='详细数据', index=False)
            self.logger.info(f"详细数据工作表: {len(df)} 行 {len(df.columns)} 列")
        else:
            # 如果数据为空，创建说明DataFrame
            self.logger.warning("数据清理后为空，创建数据说明")
            explanation_data = ['数据分析说明']
            
            if original_rows is not None and original_cols is not None:
                explanation_data.extend([
                    f'原始行数: {original_rows}',
                    f'原始列数: {original_cols}'
                ])
            
            if original_df is not None:
                explanation_data.append(f'原始列名: {list(original_df.columns)}')
            
            explanation_data.extend([
                '数据状态: 文件主要包含空值或索引列',
                '建议: 请检查原始Excel文件是否包含有效数据'
            ])
            
            df_explanation = pd.DataFrame({'数据说明': explanation_data})
            df_explanation.to_excel(writer, sheet_name='详细数据', index=False)
            self.logger.info(f"详细数据工作表: {len(df_explanation)} 行 {len(df_explanation.columns)} 列")
    
    def _write_analysis_stats_sheet(self, writer, df, source_info=None, original_df=None, original_rows=None, original_cols=None):
        """写入分析统计工作表"""
        stats_data = self._generate_analysis_stats(df, source_info, original_df, original_rows, original_cols)
        stats_df = pd.DataFrame(stats_data, columns=['统计项', '值'])
        stats_df.to_excel(writer, sheet_name='分析统计', index=False)
    
    def _generate_analysis_stats(self, df, source_info=None, original_df=None, original_rows=None, original_cols=None):
        """生成分析统计数据"""
        stats = []
        
        try:
            # 基本信息
            stats.append(['', ''])  # 空行用于格式化
            stats.append(['', ''])  # 空行用于格式化
            
            if source_info:
                stats.append(['数据来源', source_info])
            
            stats.append(['分析时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            
            # 原始数据信息（如果提供）
            if original_rows is not None and original_cols is not None:
                stats.append(['原始行数', original_rows])
                stats.append(['原始列数', original_cols])
            
            # 当前数据信息
            stats.append(['清理后行数', len(df)])
            stats.append(['清理后列数', len(df.columns)])
            
            if not df.empty:
                # 数据类型统计
                numeric_cols = len(df.select_dtypes(include=['number']).columns)
                text_cols = len(df.select_dtypes(include=['object']).columns)
                date_cols = len(df.select_dtypes(include=['datetime']).columns)
                
                stats.append(['数值列数量', numeric_cols])
                stats.append(['文本列数量', text_cols])
                stats.append(['日期列数量', date_cols])
                
                # 数据质量
                total_cells = len(df) * len(df.columns)
                empty_cells = df.isnull().sum().sum()
                
                stats.append(['空值单元格数', empty_cells])
                if total_cells > 0:
                    stats.append(['数据完整率', f"{((total_cells - empty_cells) / total_cells * 100):.1f}%"])
                else:
                    stats.append(['数据完整率', '0.0%'])
                
                # 业务统计
                self._add_business_statistics(df, stats)
                
                # 缺失值详情
                missing_info = []
                for col in df.columns:
                    missing_count = df[col].isnull().sum()
                    if missing_count > 0:
                        missing_rate = (missing_count / len(df) * 100) if len(df) > 0 else 0
                        missing_info.append(f'{col}: {missing_count}个 ({missing_rate:.1f}%)')
                
                if missing_info:
                    stats.append(['缺失值详情', '; '.join(missing_info)])
                else:
                    stats.append(['缺失值详情', '无缺失值'])
                
                # 数据类型详情
                type_info = []
                for col in df.columns:
                    dtype = str(df[col].dtype)
                    non_null_count = df[col].count()
                    type_info.append(f'{col}: {dtype} ({non_null_count}个非空值)')
                
                stats.append(['数据类型详情', '; '.join(type_info)])
                
                # 文件来源统计（批量处理特有）
                if '文件来源' in df.columns:
                    file_count = df['文件来源'].nunique()
                    stats.append(['源文件数量', file_count])
                    
                    # 显示前5个文件名
                    file_sources = df['文件来源'].unique()[:5]
                    stats.append(['源文件列表', '; '.join([str(f) for f in file_sources])])
            else:
                stats.append(['数据状态', '清理后数据为空'])
                if original_df is not None:
                    stats.append(['原始列名', str(list(original_df.columns))])
            
        except Exception as e:
            self.logger.error(f"生成统计数据失败: {str(e)}")
            stats.append(['错误信息', str(e)])
        
        return stats
    
    def _add_business_statistics(self, df, stats):
        """添加业务相关统计"""
        try:
            # 检查是否有Bug相关的列
            bug_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['bug', '缺陷', '问题', '错误'])]
            if bug_columns:
                stats.append(['Bug相关列数', len(bug_columns)])
            
            # 检查是否有级别相关的列
            level_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['级别', 'level', '等级', 'priority', '严重', 'severity'])]
            if level_columns:
                level_col = level_columns[0]
                level_counts = df[level_col].value_counts()
                for level, count in level_counts.items():
                    if pd.notna(level):
                        stats.append([f'{level}数量', count])
            
            # 检查是否有状态相关的列
            status_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['状态', 'status', '修复', 'fixed', '解决', 'resolved'])]
            if status_columns:
                status_col = status_columns[0]
                status_counts = df[status_col].value_counts()
                for status, count in status_counts.items():
                    if pd.notna(status):
                        stats.append([f'{status}数量', count])
            
            # 检查是否有类型相关的列
            type_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['类型', 'type', '分类', 'category'])]
            if type_columns:
                type_col = type_columns[0]
                type_counts = df[type_col].value_counts()
                for bug_type, count in type_counts.items():
                    if pd.notna(bug_type):
                        stats.append([f'{bug_type}数量', count])
                        
        except Exception as e:
            self.logger.error(f"添加业务统计失败: {str(e)}")