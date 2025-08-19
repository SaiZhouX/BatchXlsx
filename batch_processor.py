import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime
from excel_processor import ExcelProcessor

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BatchProcessor(ExcelProcessor):
    """批量处理Excel文件的类，继承自ExcelProcessor"""
    
    def read_all_files(self):
        """
        读取所有Excel文件
        
        Returns:
            dict: 文件名和DataFrame的字典
        """
        xlsx_files = {}
        
        # 获取所有xlsx文件路径
        file_paths = self.get_input_xlsx_files()
        
        if not file_paths:
            logger.warning("没有找到可处理的Excel文件")
            return xlsx_files
        
        # 读取每个文件
        for file_path in file_paths:
            df = self.read_excel_file(file_path)
            if not df.empty:
                xlsx_files[file_path.name] = df
        
        return xlsx_files
    
    def merge_data(self, xlsx_files):
        """
        合并所有Excel文件的数据
        
        Args:
            xlsx_files (dict): 文件名和DataFrame的字典
            
        Returns:
            pd.DataFrame: 合并后的数据
        """
        if not xlsx_files:
            logger.warning("没有数据可以合并")
            return pd.DataFrame()
            
        try:
            # 合并所有DataFrame
            merged_df = pd.concat(xlsx_files.values(), ignore_index=True)
            logger.info(f"成功合并数据，总计 {len(merged_df)} 行")
            
            return merged_df
            
        except Exception as e:
            logger.error(f"合并数据时出错: {str(e)}")
            return pd.DataFrame()
    
    def clean_dataframe(self, df):
        """
        清理DataFrame，删除空列和无用列
        
        Args:
            df (pd.DataFrame): 输入的DataFrame
            
        Returns:
            pd.DataFrame: 清理后的DataFrame
        """
        if df.empty:
            return df
        
        # 记录原始列数
        original_cols = len(df.columns)
        
        # 1. 删除所有值都为空的列
        df_cleaned = df.dropna(axis=1, how='all')
        
        # 2. 删除类似"Unnamed: X"的列
        columns_to_drop = []
        for col in df_cleaned.columns:
            if isinstance(col, str) and col.startswith('Unnamed:'):
                columns_to_drop.append(col)
        
        if columns_to_drop:
            df_cleaned = df_cleaned.drop(columns=columns_to_drop)
            logger.info(f"删除了无用列: {columns_to_drop}")
        
        # 3. 删除列名为空或只包含空白字符的列
        columns_to_drop = []
        for col in df_cleaned.columns:
            if pd.isna(col) or (isinstance(col, str) and col.strip() == ''):
                columns_to_drop.append(col)
        
        if columns_to_drop:
            df_cleaned = df_cleaned.drop(columns=columns_to_drop)
            logger.info(f"删除了空列名的列: {len(columns_to_drop)} 个")
        
        # 记录清理结果
        cleaned_cols = len(df_cleaned.columns)
        if original_cols != cleaned_cols:
            logger.info(f"列清理完成: 原有 {original_cols} 列，清理后 {cleaned_cols} 列，删除了 {original_cols - cleaned_cols} 列")
        
        return df_cleaned
    
    def generate_reports(self, merged_df):
        """
        生成各种报告
        
        Args:
            merged_df (pd.DataFrame): 合并后的数据
        """
        if merged_df.empty:
            logger.warning("没有数据可以生成报告")
            return
            
        try:
            # 生成时间戳
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # 1. 生成合并后的完整数据文件
            self.save_to_excel(merged_df, f"合并数据_{timestamp}.xlsx")
            
            # 2. 生成数据统计报告
            self._generate_summary_report(merged_df, timestamp)
            
            # 3. 生成详细分析报告
            self._generate_analysis_report(merged_df, timestamp)
            
        except Exception as e:
            logger.error(f"生成报告时出错: {str(e)}")
    
    def _generate_summary_report(self, merged_df, timestamp):
        """
        生成数据统计摘要报告
        """
        try:
            summary_file = self.output_folder / f"数据统计报告_{timestamp}.xlsx"
            
            with pd.ExcelWriter(summary_file, engine='openpyxl') as writer:
                # 基本统计信息
                basic_stats = pd.DataFrame({
                    '统计项目': ['总行数', '总列数', '文件数量', '处理时间'],
                    '数值': [
                        len(merged_df),
                        len(merged_df.columns),
                        merged_df['文件来源'].nunique() if '文件来源' in merged_df.columns else 0,
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    ]
                })
                basic_stats.to_excel(writer, sheet_name='基本统计', index=False)
                
                # 按文件来源统计
                if '文件来源' in merged_df.columns:
                    file_stats = merged_df['文件来源'].value_counts().reset_index()
                    file_stats.columns = ['文件名', '行数']
                    file_stats.to_excel(writer, sheet_name='文件统计', index=False)
                
                # 数值列统计（如果存在）
                numeric_cols = merged_df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    numeric_stats = merged_df[numeric_cols].describe()
                    numeric_stats.to_excel(writer, sheet_name='数值统计')
                
            logger.info(f"已生成统计报告: {summary_file}")
            
        except Exception as e:
            logger.error(f"生成统计报告时出错: {str(e)}")
    
    def _generate_analysis_report(self, merged_df, timestamp):
        """
        生成详细分析报告
        """
        try:
            analysis_file = self.output_folder / f"详细分析报告_{timestamp}.xlsx"
            
            # 为详细分析报告添加新列
            analysis_df = merged_df.copy()
            
            # 清理DataFrame，删除空列和无用列
            analysis_df = self.clean_dataframe(analysis_df)
            
            # 添加类型列（默认：非程序Bug）
            analysis_df['类型'] = '非程序Bug'
            
            # 添加修复状态列（默认：未修复）
            analysis_df['修复状态'] = '未修复'
            
            with pd.ExcelWriter(analysis_file, engine='openpyxl') as writer:
                # 完整数据（主工作表）
                analysis_df.to_excel(writer, sheet_name='完整数据', index=False)
                
                # 数据预览（前100行）
                preview_df = analysis_df.head(100)
                preview_df.to_excel(writer, sheet_name='数据预览', index=False)
                
                # 缺失值分析
                missing_data = pd.DataFrame({
                    '列名': analysis_df.columns,
                    '缺失值数量': analysis_df.isnull().sum().values,
                    '缺失值比例': (analysis_df.isnull().sum() / len(analysis_df) * 100).round(2).values
                })
                missing_data.to_excel(writer, sheet_name='缺失值分析', index=False)
                
                # 数据类型信息
                dtype_info = pd.DataFrame({
                    '列名': analysis_df.columns,
                    '数据类型': analysis_df.dtypes.astype(str).values,
                    '非空值数量': analysis_df.count().values
                })
                dtype_info.to_excel(writer, sheet_name='数据类型', index=False)
                
            logger.info(f"已生成分析报告: {analysis_file}")
            
        except Exception as e:
            logger.error(f"生成分析报告时出错: {str(e)}")
    
    def process(self):
        """
        执行完整的批量处理流程
        """
        logger.info("开始批量处理Excel文件...")
        
        # 1. 读取所有xlsx文件
        xlsx_files = self.read_all_files()
        
        if not xlsx_files:
            logger.error("没有找到可处理的文件，程序结束")
            return
        
        # 2. 合并数据
        merged_df = self.merge_data(xlsx_files)
        
        if merged_df.empty:
            logger.error("数据合并失败，程序结束")
            return
        
        # 3. 生成报告
        self.generate_reports(merged_df)
        
        logger.info("批量处理完成！")