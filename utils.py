"""
通用工具函数模块
"""
import re
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any
import pandas as pd

from config_manager import config
from logger_config import LoggerConfig

logger = LoggerConfig.get_logger(__name__)

class FileUtils:
    """文件操作工具类"""
    
    @staticmethod
    def is_excel_file(file_path: str) -> bool:
        """
        检查文件是否为Excel文件
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            bool: 是否为Excel文件
        """
        return config.is_supported_file(file_path)
    
    @staticmethod
    def is_temp_file(file_path: str) -> bool:
        """
        检查是否为临时文件
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            bool: 是否为临时文件
        """
        filename = Path(file_path).name
        return filename.startswith('~$') or filename.startswith('.~')
    
    @staticmethod
    def get_excel_files(folder_path: str) -> List[Path]:
        """
        获取文件夹中的所有Excel文件
        
        Args:
            folder_path (str): 文件夹路径
            
        Returns:
            List[Path]: Excel文件路径列表
        """
        folder = Path(folder_path)
        if not folder.exists():
            logger.warning(f"文件夹不存在: {folder_path}")
            return []
        
        excel_files = []
        supported_formats = config.get_supported_formats()
        
        for ext in supported_formats:
            pattern = f"*{ext}"
            files = folder.glob(pattern)
            for file_path in files:
                if not FileUtils.is_temp_file(str(file_path)):
                    excel_files.append(file_path)
        
        excel_files.sort()
        logger.info(f"在 {folder_path} 中找到 {len(excel_files)} 个Excel文件")
        return excel_files
    
    @staticmethod
    def open_file(file_path: str) -> bool:
        """
        使用系统默认程序打开文件
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            bool: 是否成功打开
        """
        try:
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):
                subprocess.call(['open', file_path])
            else:
                subprocess.call(['xdg-open', file_path])
            
            logger.info(f"已打开文件: {Path(file_path).name}")
            return True
            
        except Exception as e:
            logger.error(f"打开文件失败: {e}")
            return False
    
    @staticmethod
    def generate_timestamp_filename(prefix: str, suffix: str = '.xlsx') -> str:
        """
        生成带时间戳的文件名
        
        Args:
            prefix (str): 文件名前缀
            suffix (str): 文件扩展名
            
        Returns:
            str: 带时间戳的文件名
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"{prefix}_{timestamp}{suffix}"
    
    @staticmethod
    def find_latest_file(folder_path: str, pattern: str) -> Optional[Path]:
        """
        查找最新的匹配文件
        
        Args:
            folder_path (str): 文件夹路径
            pattern (str): 文件名模式
            
        Returns:
            Optional[Path]: 最新文件路径，如果没有找到则返回None
        """
        folder = Path(folder_path)
        if not folder.exists():
            return None
        
        matching_files = list(folder.glob(pattern))
        if not matching_files:
            return None
        
        # 按修改时间排序，返回最新的
        latest_file = max(matching_files, key=lambda x: x.stat().st_mtime)
        logger.info(f"找到最新文件: {latest_file.name}")
        return latest_file

class DataUtils:
    """数据处理工具类"""
    
    @staticmethod
    def remove_useless_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        删除无用列
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        if df.empty:
            return df
        
        original_shape = df.shape
        cleaning_config = config.get_data_cleaning_config()
        
        # 删除无用列
        if cleaning_config.get('remove_unnamed_columns', True):
            unnamed_patterns = cleaning_config.get('unnamed_column_patterns', ['Unnamed:', 'unnamed:'])
            columns_to_drop = []
            
            for col in df.columns:
                if any(pattern in str(col) for pattern in unnamed_patterns):
                    columns_to_drop.append(col)
            
            if columns_to_drop:
                df = df.drop(columns=columns_to_drop)
                logger.info(f"删除了无用列: {columns_to_drop}")
        
        cleaned_shape = df.shape
        logger.info(f"数据清理完成: 原有 {original_shape[0]} 行 {original_shape[1]} 列，"
                   f"清理后 {cleaned_shape[0]} 行 {cleaned_shape[1]} 列")
        
        return df
    
    @staticmethod
    def clean_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
        """
        删除完全为空的行
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        if df.empty:
            return df
        
        original_rows = len(df)
        
        # 删除完全为空的行
        cleaning_config = config.get_data_cleaning_config()
        if cleaning_config.get('remove_empty_rows', True):
            df = df.dropna(how='all').reset_index(drop=True)
        
        logger.info(f"  原始数据行数: {original_rows}")
        logger.info(f"  处理后数据行数: {len(df)}")
        
        if len(df) < original_rows:
            logger.info(f"  删除了 {original_rows - len(df)} 行完全为空的数据")
        
        return df
    
    @staticmethod
    def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        """
        清理DataFrame数据
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 清理后的数据
        """
        if df.empty:
            return df
        
        # 调用专门的方法清理无用列和空行
        df = DataUtils.remove_useless_columns(df)
        df = DataUtils.clean_empty_rows(df)
        
        return df
    
    @staticmethod
    def add_metadata_columns(df: pd.DataFrame, source_file: str) -> pd.DataFrame:
        """
        添加元数据列
        
        Args:
            df (pd.DataFrame): 原始数据
            source_file (str): 源文件名
            
        Returns:
            pd.DataFrame: 添加元数据后的数据
        """
        if df.empty:
            return df
        
        report_config = config.get_report_config()
        
        if report_config.get('include_source_column', True):
            df['文件来源'] = Path(source_file).name
        
        if report_config.get('include_timestamp', True):
            df['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        return df
    
    @staticmethod
    def detect_bug_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
        """
        检测Bug相关列
        
        Args:
            df (pd.DataFrame): 数据框
            
        Returns:
            Dict[str, Optional[str]]: 检测到的列名映射
        """
        bug_config = config.get_bug_analysis_config()
        
        result = {
            'bug_columns': [],
            'level_column': None,
            'source_column': None,
            'type_column': None,
            'status_column': None
        }
        
        # 检测Bug相关列
        bug_columns = bug_config.get('bug_columns', [])
        for col in df.columns:
            if any(keyword in str(col) for keyword in bug_columns):
                result['bug_columns'].append(col)
        
        # 检测级别列
        level_columns = bug_config.get('level_columns', [])
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in level_columns):
                result['level_column'] = col
                break
        
        # 检测来源列
        source_columns = bug_config.get('source_columns', [])
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in source_columns):
                result['source_column'] = col
                break
        
        # 检测类型列
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in ['类型', 'type', '种类']):
                result['type_column'] = col
                break
        
        # 检测状态列
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in ['状态', 'status', '修复']):
                result['status_column'] = col
                break
        
        return result
    
    @staticmethod
    def add_analysis_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        添加分析列：类型和修复状态
        
        Args:
            df (pd.DataFrame): 原始数据
            
        Returns:
            pd.DataFrame: 添加分析列后的数据
        """
        try:
            logger = LoggerConfig.get_logger(__name__)
            
            # 添加类型列（默认为非程序Bug）
            if '类型' not in df.columns:
                df['类型'] = '非程序Bug'
                logger.info("已添加'类型'列，默认值：非程序Bug")
            
            # 添加修复状态列（默认为未修复）
            if '修复状态' not in df.columns:
                df['修复状态'] = '未修复'
                logger.info("已添加'修复状态'列，默认值：未修复")
            else:
                # 为现有修复状态列设置默认值
                df['修复状态'] = df['修复状态'].fillna('未修复')
                logger.info("已为'修复状态'列设置默认值：未修复")
            
            return df
            
        except Exception as e:
            logger.error(f"添加分析列时出错: {e}")
            return df

class TextUtils:
    """文本处理工具类"""
    
    @staticmethod
    def extract_date_from_filename(filename: str) -> Optional[str]:
        """
        从文件名中提取日期
        
        Args:
            filename (str): 文件名
            
        Returns:
            Optional[str]: 提取的日期字符串
        """
        if pd.isna(filename) or not isinstance(filename, str):
            return None
        
        # 日期提取模式
        date_patterns = [
            r'记录(\d{4})',  # 记录后面的4位数字
            r'bug记录(\d{4})',  # bug记录后面的4位数字
            r'(\d{2})(\d{2})(?![\d])',  # 4位数字但不是年份的一部分
            r'(\d{2}/\d{2})',  # MM/DD格式
            r'(\d{2}-\d{2})',  # MM-DD格式
            r'(\d{1,2})月(\d{1,2})日?',  # 中文日期格式
            r'(\d{1,2})\.(\d{1,2})(?=\.|$)',  # 点号分隔的日期格式（如8.13）
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, filename)
            if match:
                if len(match.groups()) == 1:
                    date_str = match.group(1)
                    if len(date_str) == 4 and not date_str.startswith('20'):
                        return date_str
                elif len(match.groups()) == 2:
                    month = match.group(1).zfill(2)
                    day = match.group(2).zfill(2)
                    return f"{month}{day}"
                elif len(match.groups()) == 2:
                    month = match.group(1).zfill(2)
                    day = match.group(2).zfill(2)
                    return f"{month}{day}"
                elif '/' in match.group(1) or '-' in match.group(1):
                    return match.group(1).replace('/', '').replace('-', '')
        
        return None
    
    @staticmethod
    def extract_tester_from_filename(filename: str) -> Optional[str]:
        """
        从文件名中提取测试人姓名
        
        Args:
            filename (str): 文件名
            
        Returns:
            Optional[str]: 提取的测试人姓名
        """
        if pd.isna(filename) or not isinstance(filename, str):
            return "质检"
        
        # 常见姓名列表（优先匹配）
        common_names = ['胡先美', '王超', '李明', '张三', '李四', '王五', '赵六', '孙七']
        for name in common_names:
            if name in filename:
                return name
        
        # 正则表达式提取中文姓名
        name_patterns = [
            r'_([^_\.]+)\.xlsx?$',  # 下划线后面的姓名
            r'([一-龯]{2,4})\.xlsx?$',   # 文件名最后的中文姓名
        ]
        
        exclude_words = ['记录', '报告', '测试', '分析', '统计', '汇总', '名利场', '公司']
        
        for pattern in name_patterns:
            match = re.search(pattern, filename)
            if match:
                potential_name = match.group(1)
                if (re.search(r'[\u4e00-\u9fff]', potential_name) and 
                    len(potential_name) <= 4 and 
                    potential_name not in exclude_words):
                    return potential_name
        
        # 默认返回质检
        return "质检"
    
    @staticmethod
    def extract_date_and_tester(filename: str) -> Optional[str]:
        """
        从文件名中提取日期和测试人信息
        
        Args:
            filename (str): 文件名
            
        Returns:
            Optional[str]: 格式为"月日_姓名"的字符串
        """
        date_str = TextUtils.extract_date_from_filename(filename)
        tester_name = TextUtils.extract_tester_from_filename(filename)
        
        # 确保至少有一个有效值
        if not date_str:
            date_str = "未识别"
        if not tester_name:
            tester_name = "质检"
            
        return f"{date_str}_{tester_name}"

class ExcelUtils:
    """Excel操作工具类"""
    
    @staticmethod
    def read_excel_smart(file_path: str) -> pd.DataFrame:
        """
        智能读取Excel文件，自动检测Bug记录工作表
        
        Args:
            file_path (str): Excel文件路径
            
        Returns:
            pd.DataFrame: 读取的数据
        """
        try:
            logger.info(f"正在读取文件: {Path(file_path).name}")
            
            # 读取所有工作表
            all_sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
            
            # 查找包含Bug记录的工作表
            bug_columns = config.get_bug_columns()
            main_df = None
            
            for sheet_name, df in all_sheets.items():
                logger.info(f"  检查工作表 '{sheet_name}': {len(df)} 行 {len(df.columns)} 列")
                
                # 检查是否包含Bug记录相关的列
                matching_columns = [col for col in bug_columns if col in df.columns]
                
                if matching_columns:
                    logger.info(f"  找到Bug记录工作表: {sheet_name}, 匹配列: {matching_columns}")
                    main_df = df
                    break
            
            # 如果没找到Bug记录工作表，使用第一个工作表
            if main_df is None:
                main_df = list(all_sheets.values())[0]
                logger.info(f"  使用默认工作表，包含 {len(main_df)} 行数据")
            
            # 删除完全为空的行
            original_rows = len(main_df)
            main_df = main_df.dropna(how='all').reset_index(drop=True)
            
            logger.info(f"  原始数据行数: {original_rows}")
            logger.info(f"  处理后数据行数: {len(main_df)}")
            
            if len(main_df) < original_rows:
                logger.info(f"  删除了 {original_rows - len(main_df)} 行完全为空的数据")
            
            return main_df
            
        except Exception as e:
            logger.error(f"读取文件 {Path(file_path).name} 时出错: {e}")
            return pd.DataFrame()
    
    @staticmethod
    def save_excel_with_sheets(file_path: str, sheets_data: Dict[str, pd.DataFrame]) -> bool:
        """
        保存多个工作表到Excel文件
        
        Args:
            file_path (str): 输出文件路径
            sheets_data (Dict[str, pd.DataFrame]): 工作表数据字典
            
        Returns:
            bool: 是否保存成功
        """
        try:
            excel_config = config.get_excel_config()
            
            with pd.ExcelWriter(file_path, engine=excel_config.get('engine', 'openpyxl')) as writer:
                for sheet_name, df in sheets_data.items():
                    if not df.empty:
                        df.to_excel(writer, sheet_name=sheet_name, 
                                  index=excel_config.get('index', False))
                        logger.info(f"工作表 '{sheet_name}': {len(df)} 行 {len(df.columns)} 列")
                    else:
                        # 创建空数据说明
                        empty_df = pd.DataFrame({'说明': [f'{sheet_name}数据为空']})
                        empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        logger.warning(f"工作表 '{sheet_name}' 数据为空")
            
            logger.info(f"已保存Excel文件: {Path(file_path).name}")
            return True
            
        except Exception as e:
            logger.error(f"保存Excel文件失败: {e}")
            return False