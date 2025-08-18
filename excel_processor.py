import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime
import re

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelProcessor:
    """Excel文件处理的基类，提供基本的文件读取和处理功能"""
    
    def __init__(self, input_folder="input", output_folder="output"):
        """
        初始化Excel处理器
        
        Args:
            input_folder (str): 输入文件夹路径
            output_folder (str): 输出文件夹路径
        """
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        
        # 创建输出文件夹
        self.output_folder.mkdir(exist_ok=True)
    
    def read_excel_file(self, file_path):
        """
        读取单个Excel文件
        
        Args:
            file_path (Path): Excel文件路径
            
        Returns:
            pd.DataFrame: 读取的数据
        """
        try:
            logger.info(f"正在读取文件: {file_path.name}")
            
            # 读取所有工作表
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            
            # 查找包含Bug记录的工作表
            main_df = None
            for sheet_name, df in all_sheets.items():
                logger.info(f"  检查工作表 '{sheet_name}': {len(df)} 行, {len(df.columns)} 列")
                
                # 检查是否包含Bug记录相关的列
                bug_columns = ['编号', '严重级别', '级别', 'bug类型', '问题描述']
                matching_columns = [col for col in bug_columns if col in df.columns]
                
                if matching_columns:
                    logger.info(f"  找到Bug记录工作表: {sheet_name}, 匹配列: {matching_columns}")
                    main_df = df
                    break
                else:
                    logger.info(f"  工作表 '{sheet_name}' 不包含Bug记录列")
            
            # 如果没找到Bug记录工作表，使用第一个工作表
            if main_df is None:
                main_df = list(all_sheets.values())[0]
                logger.info(f"  使用默认工作表，包含 {len(main_df)} 行数据")
            
            # 删除完全为空的行，但保留部分有数据的行
            # 注意：不要删除只有某些列为空的行，因为这可能是有效的Bug记录
            original_rows = len(main_df)
            main_df = main_df.dropna(how='all')
            main_df = main_df.reset_index(drop=True)
            
            logger.info(f"  原始数据行数: {original_rows}")
            logger.info(f"  处理后数据行数: {len(main_df)}")
            
            if len(main_df) < original_rows:
                logger.warning(f"  删除了 {original_rows - len(main_df)} 行完全为空的数据")
            
            # 添加文件来源列
            main_df['文件来源'] = file_path.name
            main_df['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            return main_df
            
        except Exception as e:
            logger.error(f"读取文件 {file_path.name} 时出错: {str(e)}")
            return pd.DataFrame()
    
    def get_input_xlsx_files(self):
        """
        获取input文件夹里所有xlsx文件的路径
        
        Returns:
            list: xlsx文件路径列表
        """
        if not self.input_folder.exists():
            logger.error(f"input文件夹不存在: {self.input_folder}")
            return []
        
        try:
            xlsx_files = []
            for file in self.input_folder.glob("*.xlsx"):
                # 排除临时文件
                if not file.name.startswith('~$'):
                    xlsx_files.append(file)
            
            xlsx_files.sort()  # 按文件名排序
            logger.info(f"找到 {len(xlsx_files)} 个xlsx文件")
            return xlsx_files
            
        except Exception as e:
            logger.error(f"读取input文件夹时出错: {str(e)}")
            return []
    
    def save_to_excel(self, df, filename, sheet_name='Sheet1'):
        """
        保存DataFrame到Excel文件
        
        Args:
            df (pd.DataFrame): 要保存的数据
            filename (str): 文件名
            sheet_name (str): 工作表名
        """
        if df.empty:
            logger.warning(f"没有数据可以保存到 {filename}")
            return False
            
        try:
            file_path = self.output_folder / filename
            df.to_excel(file_path, sheet_name=sheet_name, index=False, engine='openpyxl')
            logger.info(f"已保存数据到: {file_path}")
            return True
            
        except Exception as e:
            logger.error(f"保存数据到 {filename} 时出错: {str(e)}")
            return False
    
    def extract_date_from_filename(self, filename):
        """
        从文件名中提取日期
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的日期字符串，如果无法提取则返回None
        """
        if pd.isna(filename):
            return None
        
        # 优先匹配月日格式，避免匹配年份
        date_patterns = [
            r'记录(\d{4})',  # 记录后面的4位数字如记录0804
            r'bug记录(\d{4})',  # bug记录后面的4位数字
            r'(\d{2})(\d{2})(?![\d])',  # 4位数字但不是年份的一部分
            r'(\d{2}/\d{2})',  # MM/DD格式
            r'(\d{2}-\d{2})',  # MM-DD格式
        ]
        
        date_str = None
        for pattern in date_patterns:
            match = re.search(pattern, str(filename))
            if match:
                if len(match.groups()) == 1:  # 单个匹配组
                    date_str = match.group(1)
                    if len(date_str) == 4 and not date_str.startswith('20'):  # 避免年份
                        break
                elif len(match.groups()) == 2:  # 两个匹配组
                    date_str = match.group(1) + match.group(2)
                    if len(date_str) == 4 and not date_str.startswith('20'):
                        break
                elif '/' in match.group(1) or '-' in match.group(1):
                    date_str = match.group(1).replace('/', '').replace('-', '')
                    break
        
        return date_str
    
    def extract_tester_from_filename(self, filename):
        """
        从文件名中提取测试人姓名
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的测试人姓名，如果无法提取则返回None
        """
        if pd.isna(filename):
            return None
        
        # 提取测试人姓名
        tester_patterns = [
            r'_([^_\.]+)\.xlsx$',  # 下划线后面的姓名，如_王超.xlsx
            r'([^_\d]+)\.xlsx$',   # 文件名最后的中文姓名
        ]
        
        tester_name = None
        for pattern in tester_patterns:
            match = re.search(pattern, str(filename))
            if match:
                potential_name = match.group(1)
                # 检查是否是中文姓名（包含中文字符且长度合理）
                if re.search(r'[\u4e00-\u9fff]', potential_name) and len(potential_name) <= 4:
                    tester_name = potential_name
                    break
        
        return tester_name