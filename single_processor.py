"""
单个文件处理器，继承自ReportGenerator
"""
from pathlib import Path

from report_generator import ReportGenerator
from config_manager import config
from logger_config import LoggerConfig
from utils import DataUtils, ExcelUtils

class SingleProcessor(ReportGenerator):
    """单个文件处理器，继承自ReportGenerator"""
    
    def __init__(self):
        super().__init__()
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
    
    def process_single_file(self, file_path):
        """
        处理单个Excel文件
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            str: 生成的报告文件路径，失败时返回None
        """
        try:
            self.logger.info(f"开始处理文件: {Path(file_path).name}")
            
            # 读取文件
            df = ExcelUtils.read_excel_smart(file_path)
            
            if df.empty:
                self.logger.error("文件读取失败或文件为空")
                return None
            
            self.logger.info(f"成功读取文件，共 {len(df)} 行数据")
            
            # 清理数据
            original_df = df.copy()
            original_rows, original_cols = df.shape
            
            df = self.clean_data(df)
            
            # 如果清理后数据为空，创建说明数据
            if df.empty or len(df.columns) == 0:
                self.logger.warning("数据清理后为空，创建数据说明")
                df = self._create_empty_data_explanation(original_df, original_rows, original_cols)
            
            # 添加分析列（类型和修复状态）
            df = DataUtils.add_analysis_columns(df)
            
            # 添加元数据
            df = DataUtils.add_metadata_columns(df, file_path)
            
            # 生成报告
            report_path = self.generate_unified_report(df, Path(file_path).stem)
            
            if report_path:
                self.logger.info(f"单个文件分析完成: {Path(file_path).name}")
                return str(report_path)
            else:
                self.logger.error("报告生成失败")
                return None
                
        except Exception as e:
            self.logger.error(f"处理单个文件时出错: {e}")
            return None
    
    def _create_empty_data_explanation(self, original_df, original_rows, original_cols):
        """
        创建空数据说明
        
        Args:
            original_df (pd.DataFrame): 原始数据
            original_rows (int): 原始行数
            original_cols (int): 原始列数
            
        Returns:
            pd.DataFrame: 说明数据
        """
        import pandas as pd
        
        explanation_data = {
            '数据说明': [
                '原始文件数据分析',
                f'原始行数: {original_rows}',
                f'原始列数: {original_cols}',
                f'原始列名: {list(original_df.columns)}',
                '数据状态: 文件主要包含空值或索引列',
                '建议: 请检查原始Excel文件是否包含有效数据'
            ]
        }
        
        return pd.DataFrame(explanation_data)