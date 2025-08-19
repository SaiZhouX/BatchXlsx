import pandas as pd
import os
from pathlib import Path
import logging
from datetime import datetime
from excel_processor import ExcelProcessor

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BugAnalyzer(ExcelProcessor):
    """Bug分析器类，继承自ExcelProcessor"""
    
    def __init__(self, input_folder="input", output_folder="output"):
        super().__init__(input_folder, output_folder)
        self.latest_report_file = None
    
    def find_latest_report(self):
        """
        查找最新的详细分析报告文件
        
        Returns:
            Path: 最新报告文件的路径，如果没有找到则返回None
        """
        try:
            # 查找所有详细分析报告文件
            report_files = list(self.output_folder.glob("详细分析报告_*.xlsx"))
            
            if not report_files:
                logger.warning("没有找到详细分析报告文件")
                return None
            
            # 按修改时间排序，获取最新的文件
            latest_file = max(report_files, key=lambda x: x.stat().st_mtime)
            self.latest_report_file = latest_file
            
            logger.info(f"找到最新报告文件: {latest_file}")
            return latest_file
            
        except Exception as e:
            logger.error(f"查找最新报告文件时出错: {str(e)}")
            return None
    
    def read_report_data(self, report_file=None):
        """
        读取报告数据
        
        Args:
            report_file (Path, optional): 报告文件路径，如果为None则使用最新的报告
            
        Returns:
            pd.DataFrame: 报告数据
        """
        if report_file is None:
            report_file = self.find_latest_report()
        
        if report_file is None:
            logger.error("没有可用的报告文件")
            return pd.DataFrame()
        
        try:
            # 读取完整数据工作表
            df = pd.read_excel(report_file, sheet_name='完整数据')
            logger.info(f"成功读取报告数据，共 {len(df)} 行")
            return df
            
        except Exception as e:
            logger.error(f"读取报告数据时出错: {str(e)}")
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
            logger.warning("没有数据可以分析")
            return pd.DataFrame()
        
        try:
            # 确保有日期列
            if '日期' not in df.columns:
                logger.error("数据中没有找到'日期'列")
                return pd.DataFrame()
            
            # 清理数据
            df_clean = df.dropna(subset=['日期'])
            
            # 按日期分组统计
            result = df_clean.groupby('日期').agg({
                '日期': 'count'  # 统计每个日期的记录数
            }).rename(columns={'日期': '总数'})
            
            # 如果数据中包含类型和修复状态列，则添加额外的统计
            if '类型' in df_clean.columns and '修复状态' in df_clean.columns:
                # 统计程序Bug数量
                program_bugs = df_clean[df_clean['类型'] == '程序Bug'].groupby('日期').size()
                result['程序Bug数'] = result.index.map(program_bugs).fillna(0).astype(int)
                
                # 统计程序Bug修复数量
                program_bugs_fixed = df_clean[(df_clean['类型'] == '程序Bug') & (df_clean['修复状态'] == '已修复')].groupby('日期').size()
                result['程序Bug修复数'] = result.index.map(program_bugs_fixed).fillna(0).astype(int)
                
                # 统计非程序Bug数量
                non_program_bugs = df_clean[df_clean['类型'] == '非程序Bug'].groupby('日期').size()
                result['非程序Bug数'] = result.index.map(non_program_bugs).fillna(0).astype(int)
                
                # 统计非程序Bug修复数量
                non_program_bugs_fixed = df_clean[(df_clean['类型'] == '非程序Bug') & (df_clean['修复状态'] == '已修复')].groupby('日期').size()
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
            
            logger.info(f"按日期分析完成，共 {len(result)} 个日期")
            return result
            
        except Exception as e:
            logger.error(f"按日期分析时出错: {str(e)}")
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
            logger.warning("没有数据可以分析")
            return pd.DataFrame()
        
        try:
            # 确保有类型列
            if '类型' not in df.columns:
                logger.warning("数据中没有找到'类型'列，使用默认分类")
                df = df.copy()
                df['类型'] = '非程序Bug'
            
            # 清理数据
            df_clean = df.dropna(subset=['类型'])
            
            # 按类型分组统计
            result = df_clean.groupby('类型').agg({
                '类型': 'count'  # 统计每个类型的记录数
            }).rename(columns={'类型': '总数'})
            
            # 如果有修复状态列，添加修复统计
            if '修复状态' in df_clean.columns:
                fixed_count = df_clean[df_clean['修复状态'] == '已修复'].groupby('类型').size()
                result['已修复数'] = result.index.map(fixed_count).fillna(0).astype(int)
                result['修复率'] = (result['已修复数'] / result['总数'] * 100).round(2)
            else:
                result['已修复数'] = 0
                result['修复率'] = 0.0
            
            # 重置索引
            result = result.reset_index()
            
            logger.info(f"按类型分析完成，共 {len(result)} 种类型")
            return result
            
        except Exception as e:
            logger.error(f"按类型分析时出错: {str(e)}")
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
                logger.error("没有数据可以生成Bug分析报告")
                return
            
            # 生成时间戳
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # 创建报告文件
            report_file = self.output_folder / f"Bug级别分析报告_{timestamp}.xlsx"
            
            with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
                # 按日期分析
                date_analysis = self.analyze_by_date(df)
                if not date_analysis.empty:
                    date_analysis.to_excel(writer, sheet_name='按日期分析', index=False)
                
                # 按类型分析
                type_analysis = self.analyze_by_type(df)
                if not type_analysis.empty:
                    type_analysis.to_excel(writer, sheet_name='按类型分析', index=False)
                
                # 原始数据（用于参考）
                df.to_excel(writer, sheet_name='原始数据', index=False)
                
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
                summary_df.to_excel(writer, sheet_name='分析汇总', index=False)
            
            logger.info(f"Bug级别分析报告已生成: {report_file}")
            
        except Exception as e:
            logger.error(f"生成Bug分析报告时出错: {str(e)}")
    
    def extract_date_and_tester_from_filename(self, filename):
        """
        从文件名中提取日期和测试人信息，返回格式为"月日_姓名"
        
        Args:
            filename (str): 文件名
            
        Returns:
            str: 提取的日期和姓名，格式为"0804_胡先美"，如果提取失败返回None
        """
        import re
        
        if pd.isna(filename) or not isinstance(filename, str):
            return None
        
        try:
            # 提取日期模式：0804, 08-04, 8月4日等
            date_patterns = [
                r'(\d{4})',  # 4位数字如0804
                r'(\d{1,2})-(\d{1,2})',  # 如08-04
                r'(\d{1,2})月(\d{1,2})日?',  # 如8月4日
                r'(\d{1,2})\.(\d{1,2})',  # 如8.4
            ]
            
            extracted_date = None
            extracted_name = None
            
            # 尝试提取日期
            for pattern in date_patterns:
                match = re.search(pattern, filename)
                if match:
                    if len(match.groups()) == 1:
                        # 4位数字格式
                        date_str = match.group(1)
                        if len(date_str) == 4:
                            extracted_date = date_str
                            break
                    elif len(match.groups()) == 2:
                        # 月-日格式
                        month = match.group(1).zfill(2)
                        day = match.group(2).zfill(2)
                        extracted_date = f"{month}{day}"
                        break
            
            # 优先查找常见姓名（精确匹配）
            common_names = ['胡先美', '王超', '李明', '张三', '李四', '王五', '赵六', '孙七']
            for name in common_names:
                if name in filename:
                    extracted_name = name
                    break
            
            # 如果没有找到常见姓名，使用正则表达式提取中文姓名
            if not extracted_name:
                # 查找文件名末尾的中文姓名（排除扩展名）
                name_pattern = r'([一-龯]{2,4})(?:\.[^.]*)?$'
                name_match = re.search(name_pattern, filename)
                if name_match:
                    potential_name = name_match.group(1)
                    # 排除一些不太可能是姓名的词汇
                    exclude_words = ['记录', '报告', '测试', '分析', '统计', '汇总', '名利场', '公司']
                    if potential_name not in exclude_words:
                        extracted_name = potential_name
            
            # 如果都提取成功，返回组合结果
            if extracted_date and extracted_name:
                return f"{extracted_date}_{extracted_name}"
            
            # 如果只有日期，尝试从文件名中找到可能的姓名
            if extracted_date:
                # 再次尝试查找姓名，使用更宽松的模式
                name_pattern = r'([一-龯]{2,3})'
                name_matches = re.findall(name_pattern, filename)
                for name in name_matches:
                    if name not in ['名利场', '记录', '测试', '报告', '分析']:
                        return f"{extracted_date}_{name}"
                
                # 如果仍然找不到人名，默认使用"质检"
                return f"{extracted_date}_质检"
            
            return None
            
        except Exception as e:
            logger.warning(f"提取文件名信息时出错: {filename}, 错误: {str(e)}")
            return None
    
    def process(self):
        """
        执行Bug分析流程
        """
        logger.info("开始Bug级别分析...")
        
        # 生成Bug分析报告
        self.generate_bug_analysis_report()
        
        logger.info("Bug级别分析完成！")