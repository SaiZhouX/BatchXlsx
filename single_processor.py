import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime
from excel_processor import ExcelProcessor
from report_generator import ReportGenerator

class SingleProcessor(ExcelProcessor, ReportGenerator):
    """单个Excel文件处理器，继承自ExcelProcessor和ReportGenerator"""
    
    def __init__(self):
        ExcelProcessor.__init__(self)
        ReportGenerator.__init__(self)
        self.logger = logging.getLogger(__name__)
    
    def process_single_file(self, file_path):
        """
        处理单个Excel文件并生成详细分析报告
        
        Args:
            file_path (str): Excel文件路径
            
        Returns:
            str: 生成的报告文件路径，如果失败返回None
        """
        try:
            self.logger.info(f"开始处理文件: {os.path.basename(file_path)}")
            
            # 读取Excel文件
            df = self.read_excel_file(Path(file_path))
            if df.empty:
                self.logger.error(f"文件读取失败或为空: {file_path}")
                return None
            
            # 保存原始数据信息
            original_df = df.copy()
            original_rows = len(df)
            original_cols = len(df.columns)
            
            self.logger.info(f"成功读取文件，共 {len(df)} 行数据")
            
            # 清理数据（使用共同的方法）
            df = self.clean_data(df)
            self.logger.info(f"数据清理完成: 原有 {original_rows} 行 {original_cols} 列，清理后 {len(df)} 行 {len(df.columns)} 列")
            
            # 生成统一格式的报告（使用共同的方法）
            base_name = Path(file_path).stem
            report_name = f"详细分析报告_{base_name}"
            source_info = os.path.basename(file_path)
            
            report_path = self.generate_unified_report(
                df=df,
                report_name=report_name,
                source_info=source_info,
                original_df=original_df,
                original_rows=original_rows,
                original_cols=original_cols
            )
            
            if report_path:
                self.logger.info(f"单个文件分析完成: {os.path.basename(file_path)}")
                return report_path
            else:
                self.logger.error("报告生成失败")
                return None
                
        except Exception as e:
            self.logger.error(f"处理文件时出错: {str(e)}")
            return None