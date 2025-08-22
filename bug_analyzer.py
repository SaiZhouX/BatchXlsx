"""
Bug分析器类，继承自ExcelProcessor
"""
import pandas as pd
from pathlib import Path
from datetime import datetime

from excel_processor import ExcelProcessor
from config_manager import config
from logger_config import LoggerConfig
from utils import FileUtils, DataUtils, TextUtils

class BugAnalyzer(ExcelProcessor):
    """Bug分析器类，继承自ExcelProcessor"""
    
    def __init__(self, input_folder=None, output_folder=None):
        super().__init__(input_folder, output_folder)
        self.latest_report_file = None
        self.logger = LoggerConfig.get_logger(self.__class__.__name__)
    
    def find_latest_report(self):
        """
        查找最新的详细分析报告文件
        
        Returns:
            Path: 最新报告文件的路径，如果没有找到则返回None
        """
        latest_file = FileUtils.find_latest_file(str(self.output_folder), "详细分析报告_*.xlsx")
        
        if latest_file:
            self.latest_report_file = latest_file
            self.logger.info(f"找到最新报告文件: {latest_file.name}")
        else:
            self.logger.warning("没有找到详细分析报告文件")
        
        return latest_file
    
    def read_report_data(self, report_file=None, sheet_name='详细数据'):
        """
        读取报告数据
        
        Args:
            report_file (Path, optional): 报告文件路径，如果为None则使用最新的报告
            sheet_name (str): 工作表名称
            
        Returns:
            pd.DataFrame: 报告数据
        """
        if report_file is None:
            report_file = self.find_latest_report()
        
        if report_file is None:
            self.logger.error("没有可用的报告文件")
            return pd.DataFrame()
        
        try:
            # 尝试读取指定工作表，如果失败则尝试其他可能的工作表名
            possible_sheets = [sheet_name, '详细数据', '完整数据', 'Sheet1']
            
            for sheet in possible_sheets:
                try:
                    df = pd.read_excel(report_file, sheet_name=sheet)
                    self.logger.info(f"成功读取报告数据，工作表: {sheet}，共 {len(df)} 行")
                    return df
                except ValueError:
                    continue
            
            # 如果所有尝试都失败，读取第一个工作表
            df = pd.read_excel(report_file)
            self.logger.info(f"使用默认工作表读取报告数据，共 {len(df)} 行")
            return df
            
        except Exception as e:
            self.logger.error(f"读取报告数据时出错: {e}")
            return pd.DataFrame()
    
    def analyze_by_date(self, df):
        """
        按日期分析Bug数据
        
        Args:
            df (pd.DataFrame): 输入数据
            
        Returns:
            pd.DataFrame: 按日期统计的结果
        """
        if df.empty:
            self.logger.warning("没有数据可以分析")
            return pd.DataFrame()
        
        try:
            # 确保有日期列
            date_column = None
            for col in df.columns:
                if any(keyword in str(col).lower() for keyword in ['日期', 'date', '时间']):
                    date_column = col
                    break
            
            if date_column is None:
                self.logger.error("数据中没有找到日期相关列")
                return pd.DataFrame()
            
            # 清理数据
            df_clean = df.dropna(subset=[date_column])
            
            # 按日期分组统计
            result = df_clean.groupby(date_column).agg({
                date_column: 'count'  # 统计每个日期的记录数
            }).rename(columns={date_column: '总数'})
            
            # 检测类型和修复状态列
            column_info = DataUtils.detect_bug_columns(df_clean)
            type_column = column_info.get('type_column')
            status_column = column_info.get('status_column')
            
            if type_column and status_column:
                # 统计程序Bug数量
                program_bugs = df_clean[df_clean[type_column] == '程序Bug'].groupby(date_column).size()
                result['程序Bug数'] = result.index.map(program_bugs).fillna(0).astype(int)
                
                # 统计程序Bug修复数量
                program_bugs_fixed = df_clean[
                    (df_clean[type_column] == '程序Bug') & 
                    (df_clean[status_column] == '已修复')
                ].groupby(date_column).size()
                result['程序Bug修复数'] = result.index.map(program_bugs_fixed).fillna(0).astype(int)
                
                # 统计非程序Bug数量
                non_program_bugs = df_clean[df_clean[type_column] == '非程序Bug'].groupby(date_column).size()
                result['非程序Bug数'] = result.index.map(non_program_bugs).fillna(0).astype(int)
                
                # 统计非程序Bug修复数量
                non_program_bugs_fixed = df_clean[
                    (df_clean[type_column] == '非程序Bug') & 
                    (df_clean[status_column] == '已修复')
                ].groupby(date_column).size()
                result['非程序Bug修复数'] = result.index.map(non_program_bugs_fixed).fillna(0).astype(int)
            else:
                # 如果没有类型和修复状态列，则填充0
                result['程序Bug数'] = 0
                result['程序Bug修复数'] = 0
                result['非程序Bug数'] = 0
                result['非程序Bug修复数'] = 0
            
            # 计算修复率
            result['程序Bug修复率'] = (result['程序Bug修复数'] / result['程序Bug数'] * 100).fillna(0).round(2)
            result['非程序Bug修复率'] = (result['非程序Bug修复数'] / result['非程序Bug数'] * 100).fillna(0).round(2)
            
            # 重置索引，使日期成为一列
            result = result.reset_index()
            
            self.logger.info(f"按日期分析完成，共 {len(result)} 个日期")
            return result
            
        except Exception as e:
            self.logger.error(f"按日期分析时出错: {e}")
            return pd.DataFrame()
    
    def analyze_by_type(self, df):
        """
        按类型分析Bug数据
        
        Args:
            df (pd.DataFrame): 输入数据
            
        Returns:
            pd.DataFrame: 按类型统计的结果
        """
        if df.empty:
            self.logger.warning("没有数据可以分析")
            return pd.DataFrame()
        
        try:
            # 检测类型列
            column_info = DataUtils.detect_bug_columns(df)
            type_column = column_info.get('type_column')
            
            if type_column is None:
                self.logger.warning("数据中没有找到类型列，使用默认分类")
                df = df.copy()
                df['类型'] = '非程序Bug'
                type_column = '类型'
            
            # 清理数据
            df_clean = df.dropna(subset=[type_column])
            
            # 按类型分组统计
            result = df_clean.groupby(type_column).agg({
                type_column: 'count'  # 统计每个类型的记录数
            }).rename(columns={type_column: '总数'})
            
            # 如果有修复状态列，添加修复统计
            status_column = column_info.get('status_column')
            if status_column:
                fixed_count = df_clean[df_clean[status_column] == '已修复'].groupby(type_column).size()
                result['已修复数'] = result.index.map(fixed_count).fillna(0).astype(int)
                result['修复率'] = (result['已修复数'] / result['总数'] * 100).round(2)
            else:
                result['已修复数'] = 0
                result['修复率'] = 0.0
            
            # 重置索引
            result = result.reset_index()
            
            self.logger.info(f"按类型分析完成，共 {len(result)} 种类型")
            return result
            
        except Exception as e:
            self.logger.error(f"按类型分析时出错: {e}")
            return pd.DataFrame()
    
    def generate_bug_analysis_report(self, df=None):
        """
        生成Bug级别分析报告
        
        Args:
            df (pd.DataFrame, optional): 输入数据，如果为None则从最新报告读取
        """
        try:
            # 如果没有提供数据，则从最新报告读取
            if df is None:
                df = self.read_report_data()
            
            if df.empty:
                self.logger.error("没有数据可以生成Bug分析报告")
                return
            
            # 生成报告文件名
            report_filename = FileUtils.generate_timestamp_filename("Bug级别分析报告")
            report_path = self.output_folder / report_filename
            
            # 准备工作表数据
            sheets_data = {}
            
            # 按日期分析
            date_analysis = self.analyze_by_date(df)
            if not date_analysis.empty:
                sheets_data['按日期分析'] = date_analysis
            
            # 按类型分析
            type_analysis = self.analyze_by_type(df)
            if not type_analysis.empty:
                sheets_data['按类型分析'] = type_analysis
            
            # 原始数据（用于参考）
            sheets_data['原始数据'] = df
            
            # 生成汇总信息
            summary_data = {
                '分析项目': [
                    '总记录数',
                    '分析日期数',
                    '类型数量',
                    '生成时间'
                ],
                '数值': [
                    len(df),
                    len(date_analysis) if not date_analysis.empty else 0,
                    len(type_analysis) if not type_analysis.empty else 0,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            sheets_data['分析汇总'] = summary_df
            
            # 保存报告
            from utils import ExcelUtils
            success = ExcelUtils.save_excel_with_sheets(str(report_path), sheets_data)
            
            if success:
                self.logger.info(f"Bug级别分析报告已生成: {report_filename}")
            else:
                self.logger.error("Bug级别分析报告生成失败")
            
        except Exception as e:
            self.logger.error(f"生成Bug分析报告时出错: {e}")
    
    def extract_date_and_tester_from_filename(self, filename):
        """
        从文件名中提取日期和测试人信息，返回格式为"月日_姓名"
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的日期和姓名，格式为"0804_胡先美"，如果提取失败返回None
        """
        result = TextUtils.extract_date_and_tester(filename)
        self.logger.debug(f"文件名 '{filename}' 提取结果: {result}")
        return result
    
    def process(self):
        """
        执行Bug分析流程
        """
        self.logger.info("开始Bug级别分析...")
        
        # 生成Bug分析报告
        self.generate_bug_analysis_report()
        
        self.logger.info("Bug级别分析完成！")