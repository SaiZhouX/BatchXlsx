"""
配置文件
"""

# 文件夹配置
INPUT_FOLDER = "input"
OUTPUT_FOLDER = "output"

# 支持的文件格式
SUPPORTED_FORMATS = [".xlsx"]

# 日志配置
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

# 报告配置
REPORT_CONFIG = {
    "include_source_column": True,  # 是否包含文件来源列
    "include_timestamp": True,      # 是否包含处理时间戳
    "preview_rows": 100,           # 数据预览行数
    "generate_charts": False,      # 是否生成图表（未实现）
}

# Excel写入配置
EXCEL_CONFIG = {
    "engine": "openpyxl",
    "index": False,
}