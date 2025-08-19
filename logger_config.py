"""
统一的日志配置模块
"""
import logging
import sys
from pathlib import Path
from datetime import datetime

class LoggerConfig:
    """日志配置管理类"""
    
    @staticmethod
    def setup_logger(name=None, level=logging.INFO, log_to_file=True, log_to_console=True):
        """
        设置统一的日志配置
        
        Args:
            name (str): 日志器名称，默认为调用模块名
            level (int): 日志级别
            log_to_file (bool): 是否输出到文件
            log_to_console (bool): 是否输出到控制台
            
        Returns:
            logging.Logger: 配置好的日志器
        """
        # 如果没有指定名称，使用调用模块的名称
        if name is None:
            frame = sys._getframe(1)
            name = frame.f_globals.get('__name__', 'unknown')
        
        logger = logging.getLogger(name)
        
        # 避免重复添加处理器
        if logger.handlers:
            return logger
        
        logger.setLevel(level)
        
        # 创建格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 控制台处理器
        if log_to_console:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(level)
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        # 文件处理器
        if log_to_file:
            # 确保logs目录存在
            log_dir = Path("logs")
            log_dir.mkdir(exist_ok=True)
            
            # 创建日志文件名（按日期）
            log_file = log_dir / f"excel_processor_{datetime.now().strftime('%Y%m%d')}.log"
            
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger

    @staticmethod
    def get_logger(name=None):
        """
        获取日志器的便捷方法
        
        Args:
            name (str): 日志器名称
            
        Returns:
            logging.Logger: 日志器实例
        """
        return LoggerConfig.setup_logger(name)