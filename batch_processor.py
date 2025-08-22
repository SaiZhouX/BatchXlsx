"""
批量文件处理器，继承自ReportGenerator
"""
import shutil
from pathlib import Path
from datetime import datetime
import pandas as pd
import re

from report_generator import ReportGenerator
from config_manager import config
from logger_config import LoggerConfig
from utils import FileUtils, DataUtils, ExcelUtils

class BatchProcessor(ReportGenerator):
    """批量文件处理器，继承自ReportGenerator"""
    
    def __init__(self, input_folder=None):
        super().__init__()
        self.input_folder = config.get_folder_path('input') if input_folder is None else Path(input_folder)
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
    
    def read_all_files(self):
        """
        读取所有Excel文件
        
        Returns:
            dict: 文件名到DataFrame的映射
        """
        xlsx_files = FileUtils.get_excel_files(str(self.input_folder))
        
        if not xlsx_files:
            self.logger.warning(f"在 {self.input_folder} 中没有找到Excel文件")
            return {}
        
        file_data = {}
        
        for file_path in xlsx_files:
            try:
                df = ExcelUtils.read_excel_smart(str(file_path))
                
                if not df.empty:
                    # 清理数据
                    df = self.clean_data(df)
                    
                    if not df.empty:
                        # 添加元数据
                        df = DataUtils.add_metadata_columns(df, str(file_path))
                        file_data[file_path.name] = df
                        self.logger.info(f"成功处理文件: {file_path.name}, {len(df)} 行数据")
                    else:
                        self.logger.warning(f"文件 {file_path.name} 清理后为空")
                else:
                    self.logger.warning(f"文件 {file_path.name} 读取为空")
                    
            except Exception as e:
                self.logger.error(f"处理文件 {file_path.name} 时出错: {e}")
        
        self.logger.info(f"成功读取 {len(file_data)} 个文件")
        return file_data
    
    def _clean_column_names(self, df):
        """
        清理列名，处理异常列名问题
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 清理列名后的数据
        """
        try:
            if df.empty:
                return df
            
            original_columns = list(df.columns)
            self.logger.info(f"原始列名: {original_columns}")
            
            # 清理列名
            new_columns = []
            for col in df.columns:
                if isinstance(col, str):
                    clean_col = col
                    
                    # 处理Unnamed列名
                    if clean_col.startswith('Unnamed:'):
                        # 尝试从第一行数据中获取有意义的列名
                        if len(df) > 0 and pd.notna(df.iloc[0][col]):
                            first_value = str(df.iloc[0][col])
                            if first_value and not first_value.isdigit():
                                clean_col = first_value[:20]  # 限制长度
                            else:
                                clean_col = f"列{len(new_columns)+1}"
                        else:
                            clean_col = f"列{len(new_columns)+1}"
                    
                    # 处理重复后缀（如严重级别_1）
                    clean_col = re.sub(r'\.\d+$', '', clean_col)  # 移除.1, .2等
                    clean_col = re.sub(r'_\d+$', '', clean_col)   # 移除_1, _2等
                    
                    # 处理特殊异常列名映射
                    # 已根据用户要求取消映射
                    column_mapping = {}
                    
                    if clean_col in column_mapping:
                        clean_col = column_mapping[clean_col]
                    
                    # 移除之前添加的额外处理逻辑，统一通过column_mapping处理
                    # 添加额外的严重级别列名处理
                    # if '严重级别' in clean_col and clean_col != '严重级别':
                    #     clean_col = '严重级别'
                    
                    new_columns.append(clean_col)
                else:
                    new_columns.append(str(col))
            
            # 新增判断：如果2列名相同，检查下每列下面是否有值，没有的话，不要对该列做处理，只保留有数据的列
            columns_to_drop = []
            
            # 在原始列名中查找重复的列
            original_column_counts = {}
            for col in original_columns:
                original_column_counts[col] = original_column_counts.get(col, 0) + 1
            
            # 对于原始列名中重复的列，检查每列下面是否有值
            for col, count in original_column_counts.items():
                if count > 1:  # 如果原始列名重复
                    # 找到所有同名列的索引
                    duplicate_indices = [i for i, c in enumerate(original_columns) if c == col]
                    
                    # 检查每列是否有数据，只保留有数据的列
                    has_data_indices = []
                    for idx in duplicate_indices:
                        # 检查该列是否有非空值
                        if not (df.iloc[:, idx].isnull().all() or (df.iloc[:, idx].astype(str).str.strip() == '').all()):
                            has_data_indices.append(idx)
                    
                    # 如果有多个有数据的列，只保留第一个
                    if has_data_indices:
                        # 删除其他有数据的列（保留第一个）
                        for idx in has_data_indices[1:]:
                            columns_to_drop.append(original_columns[idx])
                    # 如果没有有数据的列，保留第一个列（删除其他所有）
                    else:
                        for idx in duplicate_indices[1:]:
                            columns_to_drop.append(original_columns[idx])
            
            # 删除不需要的列
            if columns_to_drop:
                # 先删除列，再进行列名清理
                for col_name in columns_to_drop:
                    if col_name in df.columns:
                        df = df.drop(col_name, axis=1)
                self.logger.info(f"已删除没有数据的重复列: {columns_to_drop}")
            
            # 清理列名（在删除重复列之后）
            new_columns = []
            # 跟踪已使用的列名以避免重复
            used_column_names = set()
            for col in df.columns:
                if isinstance(col, str):
                    clean_col = col
                    
                    # 处理Unnamed列名
                    if clean_col.startswith('Unnamed:'):
                        # 尝试从第一行数据中获取有意义的列名
                        if len(df) > 0 and pd.notna(df.iloc[0][col]):
                            first_value = str(df.iloc[0][col])
                            if first_value and not first_value.isdigit():
                                clean_col = first_value[:20]  # 限制长度
                            else:
                                clean_col = f"列{len(new_columns)+1}"
                        else:
                            clean_col = f"列{len(new_columns)+1}"
                    
                    # 处理重复后缀（如严重级别_1）
                    clean_col = re.sub(r'\.\d+$', '', clean_col)  # 移除.1, .2等
                    clean_col = re.sub(r'_\d+$', '', clean_col)   # 移除_1, _2等
                    
                    # 处理特殊异常列名映射
                    # 已根据用户要求取消映射
                    column_mapping = {}
                    
                    if clean_col in column_mapping:
                        clean_col = column_mapping[clean_col]
                    
                    # 确保列名唯一性
                    original_clean_col = clean_col
                    counter = 1
                    while clean_col in used_column_names:
                        clean_col = f"{original_clean_col}.{counter}"
                        counter += 1
                    
                    used_column_names.add(clean_col)
                    new_columns.append(clean_col)
                else:
                    new_columns.append(str(col))
            
            # 应用清理后的列名
            df.columns = new_columns
            
            # 重新获取清理后的列名
            cleaned_columns = list(df.columns)
            self.logger.info(f"清理后列名: {cleaned_columns}")
            
            # 记录清理的变化
            changes = []
            for orig, clean in zip(original_columns, cleaned_columns):
                if str(orig) != str(clean):
                    changes.append(f"{orig} -> {clean}")
            
            if changes:
                self.logger.info(f"列名清理变化: {changes}")
            
            return df
            
        except Exception as e:
            self.logger.error(f"清理列名时出错: {e}")
            return df
    
    def merge_data(self, file_data_dict):
        """
        合并多个文件的数据
        
        Args:
            file_data_dict (dict): 文件名到DataFrame的映射
            
        Returns:
            pd.DataFrame: 合并后的数据
        """
        if not file_data_dict:
            self.logger.warning("没有数据可以合并")
            return pd.DataFrame()
        
        try:
            # 在合并前对每个DataFrame进行列名清理和默认值设置
            cleaned_dataframes = []
            for filename, df in file_data_dict.items():
                # 清理列名
                df = self._clean_column_names(df)
                
                # 添加分析列并设置默认值
                df = self._add_analysis_columns(df)
                
                # 重置索引以确保没有重复索引问题
                df = df.reset_index(drop=True)
                
                cleaned_dataframes.append(df)
            
            # 使用concat合并，忽略索引
            merged_df = pd.concat(cleaned_dataframes, ignore_index=True, sort=False)
            
            self.logger.info(f"成功合并数据，总计 {len(merged_df)} 行")
            
            # 记录每个文件的数据行数
            for filename, df in file_data_dict.items():
                self.logger.info(f"  {filename}: {len(df)} 行")
            
            return merged_df
            
        except Exception as e:
            self.logger.error(f"合并数据时出错: {e}")
            return pd.DataFrame()
    
    def _add_analysis_columns(self, df):
        """
        添加分析列：类型和修复状态
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 添加分析列后的数据
        """
        try:
            # 添加类型列（默认为非程序Bug）
            if '类型' not in df.columns:
                df['类型'] = '非程序Bug'
                self.logger.info("已添加'类型'列，默认值：非程序Bug")
            
            # 添加修复状态列（默认为未修复）
            if '修复状态' not in df.columns:
                df['修复状态'] = '未修复'
                self.logger.info("已添加'修复状态'列，默认值：未修复")
            else:
                # 为现有修复状态列设置默认值
                df['修复状态'] = df['修复状态'].fillna('未修复')
                self.logger.info("已为'修复状态'列设置默认值：未修复")
            
            # 移除智能判断类型（基于关键词）
            # df = self._infer_bug_type(df)
            
            # 移除智能判断修复状态（基于关键词）
            # df = self._infer_fix_status(df)
            
            return df
            
        except Exception as e:
            self.logger.error(f"添加分析列时出错: {e}")
            return df
    
    def generate_reports(self, merged_df, processed_files=None):
        """
        生成各种报告
        
        Args:
            merged_df (pd.DataFrame): 合并后的数据
            processed_files (list): 处理的文件列表
        """
        if merged_df.empty:
            self.logger.warning("没有数据可以生成报告")
            return
        
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # 1. 保存合并数据（详细分析报告）
            merged_filename = f"详细分析报告_{timestamp}.xlsx"
            merged_path = self.output_folder / merged_filename
            
            if ExcelUtils.save_excel_with_sheets(str(merged_path), {'详细分析报告': merged_df}):
                self.logger.info(f"已保存详细分析报告到: {merged_filename}")
            
            # 2. 生成统计报告
            stats_filename = f"数据统计报告_{timestamp}.xlsx"
            stats_path = self.output_folder / stats_filename
            
            stats_data = self._generate_statistics_report(merged_df, processed_files)
            if ExcelUtils.save_excel_with_sheets(str(stats_path), {'统计数据': stats_data}):
                self.logger.info(f"已生成统计报告: {stats_filename}")
            
            # 3. 生成统一分析报告
            unified_report_path = self.generate_unified_report(merged_df, f"批量分析_{timestamp}")
            
            if unified_report_path:
                self.logger.info(f"已生成统一分析报告: {unified_report_path.name}")
            
        except Exception as e:
            self.logger.error(f"生成报告时出错: {e}")
    
    def _generate_statistics_report(self, merged_df, processed_files):
        """
        生成统计报告数据
        
        Args:
            merged_df (pd.DataFrame): 合并后的数据
            processed_files (list): 处理的文件列表
            
        Returns:
            pd.DataFrame: 统计数据
        """
        try:
            stats_data = []
            
            # 基本统计
            stats_data.append(['处理时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            stats_data.append(['处理文件数', len(processed_files) if processed_files else 0])
            stats_data.append(['总数据行数', len(merged_df)])
            stats_data.append(['总数据列数', len(merged_df.columns)])
            
            # 文件来源统计
            if '文件来源' in merged_df.columns:
                source_counts = merged_df['文件来源'].value_counts()
                stats_data.append(['', ''])  # 空行分隔
                stats_data.append(['文件来源统计', ''])
                for source, count in source_counts.items():
                    stats_data.append([f'  {source}', count])
            
            # 数据质量统计
            stats_data.append(['', ''])  # 空行分隔
            stats_data.append(['数据质量统计', ''])
            
            # 缺失值统计
            missing_counts = merged_df.isnull().sum()
            total_missing = missing_counts.sum()
            stats_data.append(['总缺失值数', total_missing])
            
            if total_missing > 0:
                stats_data.append(['缺失值详情', ''])
                for col, missing_count in missing_counts.items():
                    if missing_count > 0:
                        missing_rate = (missing_count / len(merged_df)) * 100
                        stats_data.append([f'  {col}', f'{missing_count} ({missing_rate:.1f}%)'])
            
            # 重复值统计
            duplicate_count = merged_df.duplicated().sum()
            stats_data.append(['重复行数', duplicate_count])
            
            # 新增列统计
            if '类型' in merged_df.columns:
                stats_data.append(['', ''])  # 空行分隔
                stats_data.append(['Bug类型统计', ''])
                type_counts = merged_df['类型'].value_counts()
                for bug_type, count in type_counts.items():
                    percentage = (count / len(merged_df)) * 100
                    stats_data.append([f'  {bug_type}', f'{count} ({percentage:.1f}%)'])
            
            if '修复状态' in merged_df.columns:
                stats_data.append(['', ''])  # 空行分隔
                stats_data.append(['修复状态统计', ''])
                status_counts = merged_df['修复状态'].value_counts()
                for status, count in status_counts.items():
                    percentage = (count / len(merged_df)) * 100
                    stats_data.append([f'  {status}', f'{count} ({percentage:.1f}%)'])
                
                # 计算修复率
                if '已修复' in status_counts:
                    total_bugs = len(merged_df)
                    fix_rate = (status_counts['已修复'] / total_bugs) * 100
                    stats_data.append(['总体修复率', f'{fix_rate:.1f}%'])
            
            return pd.DataFrame(stats_data, columns=['统计项', '值'])
            
        except Exception as e:
            self.logger.error(f"生成统计报告数据时出错: {e}")
            return pd.DataFrame({'统计项': ['错误'], '值': [str(e)]})
    
    def process_batch(self):
        """
        执行批量处理流程
        
        Returns:
            bool: 处理是否成功
        """
        try:
            self.logger.info("开始批量处理...")
            
            # 1. 读取所有文件
            file_data = self.read_all_files()
            
            if not file_data:
                self.logger.error("没有找到可处理的文件")
                return False
            
            # 2. 合并数据
            merged_df = self.merge_data(file_data)
            
            if merged_df.empty:
                self.logger.error("数据合并失败")
                return False
            
            # 3. 生成报告
            processed_files = list(file_data.keys())
            self.generate_reports(merged_df, processed_files)
            
            self.logger.info("批量处理完成！")
            return True
            
        except Exception as e:
            self.logger.error(f"批量处理过程中出错: {e}")
            return False