"""
批量文件处理器，继承自ReportGenerator
"""
import shutil
from pathlib import Path
from datetime import datetime

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
            import pandas as pd
            return pd.DataFrame()
        
        try:
            # 合并所有数据
            all_dataframes = list(file_data_dict.values())
            
            # 使用concat合并，忽略索引
            import pandas as pd
            merged_df = pd.concat(all_dataframes, ignore_index=True, sort=False)
            
            self.logger.info(f"成功合并数据，总计 {len(merged_df)} 行")
            
            # 记录每个文件的数据行数
            for filename, df in file_data_dict.items():
                self.logger.info(f"  {filename}: {len(df)} 行")
            
            return merged_df
            
        except Exception as e:
            self.logger.error(f"合并数据时出错: {e}")
            import pandas as pd
            return pd.DataFrame()
    
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
            
            # 1. 保存合并数据
            merged_filename = f"合并数据_{timestamp}.xlsx"
            merged_path = self.output_folder / merged_filename
            
            if ExcelUtils.save_excel_with_sheets(str(merged_path), {'合并数据': merged_df}):
                self.logger.info(f"已保存数据到: {merged_filename}")
            
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
        import pandas as pd
        
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