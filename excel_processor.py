"""
Excel文件处理的基类，提供基本的文件读取和处理功能
"""
import pandas as pd
from pathlib import Path
from datetime import datetime

from config_manager import config
from logger_config import LoggerConfig
from utils import FileUtils, DataUtils, TextUtils, ExcelUtils

class ExcelProcessor:
    """Excel文件处理的基类，提供基本的文件读取和处理功能"""
    
    def __init__(self, input_folder=None, output_folder=None):
        """
        初始化Excel处理器
        
        Args:
            input_folder (str): 输入文件夹路径
            output_folder (str): 输出文件夹路径
        """
        self.input_folder = config.get_folder_path('input') if input_folder is None else Path(input_folder)
        self.output_folder = config.get_folder_path('output') if output_folder is None else Path(output_folder)
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
    
    def read_excel_file(self, file_path):
        """
        读取单个Excel文件
        
        Args:
            file_path (Path): Excel文件路径
            
        Returns:
            pd.DataFrame: 读取的数据
        """
        df = ExcelUtils.read_excel_smart(str(file_path))
        
        if not df.empty:
            # 添加元数据列
            df = DataUtils.add_metadata_columns(df, str(file_path))
        
        return df
    
    def get_input_xlsx_files(self):
        """
        获取input文件夹里所有xlsx文件的路径
        
        Returns:
            list: xlsx文件路径列表
        """
        return FileUtils.get_excel_files(str(self.input_folder))
    
    def save_to_excel(self, df, filename, sheet_name='Sheet1'):
        """
        保存DataFrame到Excel文件
        
        Args:
            df (pd.DataFrame): 要保存的数据
            filename (str): 文件名
            sheet_name (str): 工作表名
            
        Returns:
            bool: 是否保存成功
        """
        if df.empty:
            self.logger.warning(f"没有数据可以保存到 {filename}")
            return False
        
        file_path = self.output_folder / filename
        sheets_data = {sheet_name: df}
        return ExcelUtils.save_excel_with_sheets(str(file_path), sheets_data)
    
    def extract_date_from_filename(self, filename):
        """
        从文件名中提取日期
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的日期字符串，如果无法提取则返回None
        """
        return TextUtils.extract_date_from_filename(filename)
    
    def extract_tester_from_filename(self, filename):
        """
        从文件名中提取测试人姓名
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的测试人姓名，如果无法提取则返回None
        """
        return TextUtils.extract_tester_from_filename(filename)