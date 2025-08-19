"""
统一的配置管理模块
"""
import os
from pathlib import Path
from typing import Dict, Any, List

class ConfigManager:
    """配置管理类，统一管理所有配置项"""
    
    # 默认配置
    DEFAULT_CONFIG = {
        # 文件夹配置
        'folders': {
            'input': 'input',
            'output': 'output',
            'temp': 'temp_input',
            'logs': 'logs'
        },
        
        # 文件格式配置
        'file_formats': {
            'supported_excel': ['.xlsx', '.xls'],
            'output_format': '.xlsx'
        },
        
        # Excel处理配置
        'excel': {
            'engine': 'openpyxl',
            'index': False,
            'preview_rows': 100,
            'max_file_size_mb': 100
        },
        
        # 报告配置
        'report': {
            'include_source_column': True,
            'include_timestamp': True,
            'generate_charts': False,
            'auto_open_report': True
        },
        
        # Bug分析配置
        'bug_analysis': {
            'level_mapping': {
                'S-严重': 'S级',
                'A-重要': 'A级',
                'B-一般': 'B级',
                'C-轻微': 'C级'
            },
            'default_level': '未分级',
            'bug_columns': ['编号', '严重级别', '级别', 'bug类型', '问题描述'],
            'level_columns': ['级别', 'level', '等级', 'priority', '严重', 'severity'],
            'source_columns': ['来源', 'source', '文件', 'file', '文件来源']
        },
        
        # 数据清理配置
        'data_cleaning': {
            'remove_unnamed_columns': True,
            'remove_empty_rows': True,
            'fill_missing_values': False,
            'unnamed_column_patterns': ['Unnamed:', 'unnamed:']
        },
        
        # 日志配置
        'logging': {
            'level': 'INFO',
            'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            'log_to_file': True,
            'log_to_console': True
        },
        
        # GUI配置
        'gui': {
            'window_size': '1000x700',
            'theme': 'winnative',
            'auto_scroll_log': True,
            'show_progress': True
        }
    }
    
    def __init__(self):
        """初始化配置管理器"""
        self._config = self.DEFAULT_CONFIG.copy()
        self._ensure_directories()
    
    def _ensure_directories(self):
        """确保必要的目录存在"""
        for folder_key, folder_path in self._config['folders'].items():
            Path(folder_path).mkdir(exist_ok=True)
    
    def get(self, key_path: str, default=None) -> Any:
        """
        获取配置值，支持点号分隔的路径
        
        Args:
            key_path (str): 配置键路径，如 'folders.input'
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        keys = key_path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any):
        """
        设置配置值
        
        Args:
            key_path (str): 配置键路径
            value (Any): 配置值
        """
        keys = key_path.split('.')
        config = self._config
        
        # 导航到最后一级的父级
        for key in keys[:-1]:
            if key not in config:
                config[key] = {}
            config = config[key]
        
        # 设置最终值
        config[keys[-1]] = value
    
    def get_folder_path(self, folder_name: str) -> Path:
        """
        获取文件夹路径
        
        Args:
            folder_name (str): 文件夹名称
            
        Returns:
            Path: 文件夹路径对象
        """
        folder_path = self.get(f'folders.{folder_name}', folder_name)
        path = Path(folder_path)
        path.mkdir(exist_ok=True)
        return path
    
    def get_supported_formats(self) -> List[str]:
        """获取支持的文件格式列表"""
        return self.get('file_formats.supported_excel', ['.xlsx'])
    
    def get_excel_config(self) -> Dict[str, Any]:
        """获取Excel处理配置"""
        return self.get('excel', {})
    
    def get_report_config(self) -> Dict[str, Any]:
        """获取报告配置"""
        return self.get('report', {})
    
    def get_bug_analysis_config(self) -> Dict[str, Any]:
        """获取Bug分析配置"""
        return self.get('bug_analysis', {})
    
    def get_data_cleaning_config(self) -> Dict[str, Any]:
        """获取数据清理配置"""
        return self.get('data_cleaning', {})
    
    def get_logging_config(self) -> Dict[str, Any]:
        """获取日志配置"""
        return self.get('logging', {})
    
    def get_gui_config(self) -> Dict[str, Any]:
        """获取GUI配置"""
        return self.get('gui', {})
    
    def is_supported_file(self, file_path: str) -> bool:
        """
        检查文件是否为支持的格式
        
        Args:
            file_path (str): 文件路径
            
        Returns:
            bool: 是否支持
        """
        file_ext = Path(file_path).suffix.lower()
        return file_ext in self.get_supported_formats()
    
    def get_level_mapping(self) -> Dict[str, str]:
        """获取Bug级别映射"""
        return self.get('bug_analysis.level_mapping', {})
    
    def get_bug_columns(self) -> List[str]:
        """获取Bug相关列名"""
        return self.get('bug_analysis.bug_columns', [])
    
    def get_level_columns(self) -> List[str]:
        """获取级别相关列名"""
        return self.get('bug_analysis.level_columns', [])
    
    def get_source_columns(self) -> List[str]:
        """获取来源相关列名"""
        return self.get('bug_analysis.source_columns', [])

# 全局配置实例
config = ConfigManager()