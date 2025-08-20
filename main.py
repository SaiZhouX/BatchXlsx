import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import shutil
import threading
from pathlib import Path
import pandas as pd
from datetime import datetime
import logging

# 导入自定义模块
from batch_processor import BatchProcessor
from bug_analyzer import BugAnalyzer
from data_validator import DataValidator
from single_processor import SingleProcessor

class ExcelAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel批量处理与Bug分析工具")
        self.root.geometry("1000x700")
        
        # 存储文件列表
        self.file_list = []
        # 标记是否为单个文件模式
        self.single_file_mode = False
        
        # 创建界面
        self.create_widgets()
        
        # 配置日志
        self.setup_logging()
    
    def setup_logging(self):
        """配置日志"""
        # 创建日志处理器，将日志输出到GUI
        self.log_handler = GUILogHandler(self.log_text)
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[self.log_handler]
        )
        self.logger = logging.getLogger(__name__)
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel批量处理与Bug分析工具", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 添加单个文件按钮
        ttk.Button(file_frame, text="添加单个Excel文件", 
                  command=self.add_single_file).grid(row=0, column=0, padx=(0, 10))
        
        # 添加文件夹按钮
        ttk.Button(file_frame, text="添加文件夹", 
                  command=self.add_folder).grid(row=0, column=1, padx=(0, 10))
        
        # 清空文件列表按钮
        ttk.Button(file_frame, text="清空列表", 
                  command=self.clear_files).grid(row=0, column=2)
        
        # 文件列表显示
        list_frame = ttk.LabelFrame(main_frame, text="已选择的文件", padding="10")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建文件列表
        self.file_listbox = tk.Listbox(list_frame, height=8)
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        # 开始分析按钮
        self.analyze_button = ttk.Button(button_frame, text="开始分析", 
                                        command=self.start_analysis, 
                                        style='Accent.TButton')
        self.analyze_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Bug级别分析按钮
        self.bug_analysis_button = ttk.Button(button_frame, text="Bug级别分析", 
                                             command=self.start_bug_analysis)
        self.bug_analysis_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 进度条
        self.progress = ttk.Progressbar(button_frame, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 创建Notebook用于显示不同的结果
        self.notebook = ttk.Notebook(result_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Bug统计标签页
        self.bug_stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.bug_stats_frame, text="Bug级别统计")
        
        # 创建Bug统计表格
        self.create_bug_stats_table()
        
        # 日志标签页
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="处理日志")
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def create_bug_stats_table(self):
        """创建Bug统计表格"""
        # 表格框架
        table_frame = ttk.Frame(self.bug_stats_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建Treeview表格 - 调整列顺序，将统计列移到文件名称后面
        columns = ('文件名称', '总计', '程序Bug数', '程序Bug修复数', '非程序Bug数', '非程序Bug修复数', 'S级', 'A级', 'B级', 'C级', '未分级')
        self.bug_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # 设置列标题和列宽
        column_widths = {
            '文件名称': 100,
            '总计': 60,
            '程序Bug数': 80,
            '程序Bug修复数': 90,
            '非程序Bug数': 90,
            '非程序Bug修复数': 100,
            'S级': 50,
            'A级': 50,
            'B级': 50,
            'C级': 50,
            '未分级': 60
        }
        
        for col in columns:
            self.bug_tree.heading(col, text=col)
            width = column_widths.get(col, 60)
            self.bug_tree.column(col, width=width, anchor='center')
        
        # 添加滚动条
        v_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.bug_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient="horizontal", command=self.bug_tree.xview)
        
        self.bug_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # 布局
        self.bug_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # 统计摘要标签
        self.summary_label = ttk.Label(self.bug_stats_frame, text="", font=('Arial', 10))
        self.summary_label.pack(pady=10)
    
    def add_single_file(self):
        """添加单个Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            if file_path not in self.file_list:
                self.file_list.append(file_path)
                self.file_listbox.insert(tk.END, os.path.basename(file_path))
                self.log_message(f"已添加文件: {os.path.basename(file_path)}")
            else:
                messagebox.showwarning("警告", "文件已存在于列表中")
    
    def add_folder(self):
        """添加文件夹中的所有Excel文件"""
        folder_path = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        
        if folder_path:
            excel_files = []
            for ext in ['*.xlsx', '*.xls']:
                excel_files.extend(Path(folder_path).glob(ext))
            
            added_count = 0
            for file_path in excel_files:
                file_str = str(file_path)
                if file_str not in self.file_list and not file_path.name.startswith('~$'):
                    self.file_list.append(file_str)
                    self.file_listbox.insert(tk.END, file_path.name)
                    added_count += 1
            
            if added_count > 0:
                self.log_message(f"从文件夹 {os.path.basename(folder_path)} 添加了 {added_count} 个Excel文件")
            else:
                messagebox.showinfo("信息", "文件夹中没有找到新的Excel文件")
    
    def clear_files(self):
        """清空文件列表"""
        self.file_list.clear()
        self.file_listbox.delete(0, tk.END)
        self.clear_bug_stats()
        self.log_message("已清空文件列表")
    
    def clear_bug_stats(self):
        """清空Bug统计表格"""
        for item in self.bug_tree.get_children():
            self.bug_tree.delete(item)
        self.summary_label.config(text="")
    
    def log_message(self, message):
        """在日志区域显示消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def start_analysis(self):
        """开始分析"""
        if not self.file_list:
            messagebox.showwarning("警告", "请先添加要分析的Excel文件")
            return
        
        # 禁用分析按钮，启动进度条
        self.analyze_button.config(state='disabled')
        self.progress.start()
        
        # 清空之前的结果
        self.clear_bug_stats()
        
        # 自动识别处理模式：只有1个文件时使用单个文件处理，否则使用批量处理
        if len(self.file_list) == 1:
            self.single_file_mode = True
            self.log_message("检测到单个文件，使用单个文件处理模式")
            # 在新线程中执行单个文件分析
            analysis_thread = threading.Thread(target=self.perform_single_analysis)
            analysis_thread.daemon = True
            analysis_thread.start()
        else:
            self.single_file_mode = False
            self.log_message(f"检测到{len(self.file_list)}个文件，自动使用批量处理模式")
            # 在新线程中执行批量分析
            analysis_thread = threading.Thread(target=self.perform_analysis)
            analysis_thread.daemon = True
            analysis_thread.start()
    
    def perform_single_analysis(self):
        """执行单个文件分析（在后台线程中运行）"""
        try:
            self.log_message("开始分析单个Excel文件...")
            
            # 获取单个文件路径
            file_path = self.file_list[0]
            
            # 创建单个文件处理器
            processor = SingleProcessor()
            
            # 处理单个文件
            report_path = processor.process_single_file(file_path)
            
            if not report_path:
                self.root.after(0, lambda: self.analysis_complete("单个文件分析失败"))
                return
            
            # 在主线程中打开详细分析报告
            self.root.after(0, lambda file=str(report_path): self.open_report_file(file))
            
            self.root.after(0, lambda: self.analysis_complete("单个文件分析完成！"))
            
        except Exception as e:
            error_msg = f"单个文件分析过程中出现错误: {str(e)}"
            self.root.after(0, lambda: self.analysis_complete(error_msg))

    def perform_analysis(self):
        """执行批量分析（在后台线程中运行）"""
        try:
            self.log_message("开始分析Excel文件...")
            
            # 创建临时input文件夹
            temp_input = Path("temp_input")
            temp_input.mkdir(exist_ok=True)
            
            # 复制文件到临时文件夹
            for file_path in self.file_list:
                src_path = Path(file_path)
                dst_path = temp_input / src_path.name
                
                # 如果目标文件不存在，则复制
                if not dst_path.exists():
                    shutil.copy2(src_path, dst_path)
            
            # 执行批量处理
            processor = BatchProcessor(input_folder=str(temp_input))
            xlsx_files = processor.read_all_files()
            
            if not xlsx_files:
                self.root.after(0, lambda: self.analysis_complete("没有找到可处理的文件"))
                return
            
            merged_df = processor.merge_data(xlsx_files)
            
            if merged_df.empty:
                self.root.after(0, lambda: self.analysis_complete("数据合并失败"))
                return
            
            # 生成报告
            processor.generate_reports(merged_df)
            
            # 查找最新的详细分析报告并自动打开
            validator = DataValidator()
            latest_file = validator.find_latest_report("详细分析报告")
            
            if latest_file:
                # 在主线程中打开详细分析报告
                self.root.after(0, lambda file=str(latest_file): self.open_report_file(file))
            
            # 清理临时文件夹
            shutil.rmtree(temp_input, ignore_errors=True)
            
            self.root.after(0, lambda: self.analysis_complete("分析完成！"))
            
        except Exception as e:
            error_msg = f"分析过程中出现错误: {str(e)}"
            self.root.after(0, lambda: self.analysis_complete(error_msg))
    
    def analyze_bug_levels_for_gui(self, excel_file_path):
        """为GUI优化的Bug级别分析"""
        try:
            df = pd.read_excel(excel_file_path)
            
            self.log_message(f"读取到 {len(df)} 行数据")
            
            # 检查来源文件列和级别相关的列
            source_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['来源', 'source', '文件', 'file'])]
            level_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['级别', 'level', '等级', 'priority', '严重', 'severity'])]
            
            self.log_message(f"检测到来源列: {source_columns}")
            self.log_message(f"检测到级别列: {level_columns}")
            
            if not source_columns or not level_columns:
                self.log_message("未找到必要的列")
                return None
            
            source_col = source_columns[0]
            level_col = level_columns[0]
            
            # 数据预处理 - 只删除来源列为空的行
            self.log_message(f"数据预处理前行数: {len(df)}")
            df_clean = df.dropna(subset=[source_col]).copy()
            self.log_message(f"删除来源列为空后行数: {len(df_clean)}")
            
            # 对级别为空的数据给默认值
            df_clean[level_col] = df_clean[level_col].fillna('未分级')
            
            # 从来源文件名中提取日期和测试人信息
            analyzer = BugAnalyzer()
            df_clean['文件名称'] = df_clean[source_col].apply(analyzer.extract_date_and_tester_from_filename)
            
            # 统计文件名称提取成功的数量
            filename_extracted = df_clean['文件名称'].notna().sum()
            self.log_message(f"成功提取文件名称的行数: {filename_extracted}")
            
            # 过滤掉无法提取文件名称的记录
            df_clean = df_clean.dropna(subset=['文件名称'])
            self.log_message(f"过滤无文件名称记录后行数: {len(df_clean)}")
            
            # 处理级别名称，统一格式
            level_mapping = {
                'S-严重': 'S级',
                'A-重要': 'A级', 
                'B-一般': 'B级',
                'C-轻微': 'C级'
            }
            df_clean['级别'] = df_clean[level_col].map(level_mapping).fillna(df_clean[level_col])
            
            # 统计各级别Bug数量
            result = df_clean.groupby(['文件名称', '级别']).size().unstack(fill_value=0)
            
            # 确保包含所有级别的列
            for level in ['S级', 'A级', 'B级', 'C级', '未分级']:
                if level not in result.columns:
                    result[level] = 0
            
            # 重新排序列
            level_order = ['S级', 'A级', 'B级', 'C级', '未分级']
            available_levels = [level for level in level_order if level in result.columns]
            other_levels = [col for col in result.columns if col not in level_order]
            
            result = result[available_levels + other_levels]
            
            # 添加总计列
            result['总计'] = result.sum(axis=1)
            
            # 如果数据中包含类型和修复状态列，则添加额外的统计
            if '类型' in df_clean.columns and '修复状态' in df_clean.columns:
                # 统计程序Bug数量
                program_bugs = df_clean[df_clean['类型'] == '程序Bug'].groupby('文件名称').size()
                result['程序Bug数'] = result.index.map(program_bugs).fillna(0).astype(int)
                
                # 统计程序Bug修复数量
                program_bugs_fixed = df_clean[(df_clean['类型'] == '程序Bug') & (df_clean['修复状态'] == '已修复')].groupby('文件名称').size()
                result['程序Bug修复数'] = result.index.map(program_bugs_fixed).fillna(0).astype(int)
                
                # 统计非程序Bug数量
                non_program_bugs = df_clean[df_clean['类型'] == '非程序Bug'].groupby('文件名称').size()
                result['非程序Bug数'] = result.index.map(non_program_bugs).fillna(0).astype(int)
                
                # 统计非程序Bug修复数量
                non_program_bugs_fixed = df_clean[(df_clean['类型'] == '非程序Bug') & (df_clean['修复状态'] == '已修复')].groupby('文件名称').size()
                result['非程序Bug修复数'] = result.index.map(non_program_bugs_fixed).fillna(0).astype(int)
            else:
                # 如果没有类型和修复状态列，则填充0
                result['程序Bug数'] = 0
                result['程序Bug修复数'] = 0
                result['非程序Bug数'] = 0
                result['非程序Bug修复数'] = 0
            
            return result
            
        except Exception as e:
            self.log_message(f"Bug级别分析出错: {str(e)}")
            return None
    
    def update_bug_stats(self, bug_stats):
        """更新Bug统计表格"""
        if bug_stats is None:
            self.log_message("无法生成Bug统计数据")
            return
        
        # 清空现有数据
        self.clear_bug_stats()
        
        # 添加数据到表格
        total_bugs = 0
        level_totals = {}
        
        for filename, row in bug_stats.iterrows():
            # 按新的列顺序填充数据：文件名称, 总计, 程序Bug数, 程序Bug修复数, 非程序Bug数, 非程序Bug修复数, S级, A级, B级, C级, 未分级
            values = [
                filename,
                str(row.get('总计', 0)),
                str(row.get('程序Bug数', 0)),
                str(row.get('程序Bug修复数', 0)),
                str(row.get('非程序Bug数', 0)),
                str(row.get('非程序Bug修复数', 0)),
                str(row.get('S级', 0)),
                str(row.get('A级', 0)),
                str(row.get('B级', 0)),
                str(row.get('C级', 0)),
                str(row.get('未分级', 0))
            ]
            
            # 统计级别总数（不包括统计列）
            for col in ['S级', 'A级', 'B级', 'C级', '未分级']:
                value = row.get(col, 0)
                level_totals[col] = level_totals.get(col, 0) + value
            
            total_bugs += row.get('总计', 0)
            self.bug_tree.insert('', 'end', values=values)
        
        # 添加总计行
        if len(bug_stats) > 1:
            # 计算程序Bug和非程序Bug的总计
            total_program_bugs = bug_stats['程序Bug数'].sum() if '程序Bug数' in bug_stats.columns else 0
            total_program_bugs_fixed = bug_stats['程序Bug修复数'].sum() if '程序Bug修复数' in bug_stats.columns else 0
            total_non_program_bugs = bug_stats['非程序Bug数'].sum() if '非程序Bug数' in bug_stats.columns else 0
            total_non_program_bugs_fixed = bug_stats['非程序Bug修复数'].sum() if '非程序Bug修复数' in bug_stats.columns else 0
            
            # 按新的列顺序填充总计行：文件名称, 总计, 程序Bug数, 程序Bug修复数, 非程序Bug数, 非程序Bug修复数, S级, A级, B级, C级, 未分级
            total_row = [
                '总计',
                str(total_bugs),
                str(total_program_bugs),
                str(total_program_bugs_fixed),
                str(total_non_program_bugs),
                str(total_non_program_bugs_fixed),
                str(level_totals.get('S级', 0)),
                str(level_totals.get('A级', 0)),
                str(level_totals.get('B级', 0)),
                str(level_totals.get('C级', 0)),
                str(level_totals.get('未分级', 0))
            ]
            
            self.bug_tree.insert('', 'end', values=total_row, tags=('total',))
            self.bug_tree.tag_configure('total', background='lightgray', font=('Arial', 9, 'bold'))
        
        # 更新统计摘要
        summary_text = f"总Bug数: {total_bugs}  |  "
        for level in ['S级', 'A级', 'B级', 'C级', '未分级']:
            count = level_totals.get(level, 0)
            if count > 0:
                summary_text += f"{level}: {count}  "
        
        self.summary_label.config(text=summary_text)
        
        self.log_message(f"Bug统计完成，共发现 {total_bugs} 个Bug")
    
    def open_report_file(self, file_path):
        """打开报告文件"""
        try:
            import subprocess
            import sys
            
            if sys.platform.startswith('win'):
                # Windows
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):
                # macOS
                subprocess.call(['open', file_path])
            else:
                # Linux
                subprocess.call(['xdg-open', file_path])
            
            self.log_message(f"已打开详细分析报告: {os.path.basename(file_path)}")
            
        except Exception as e:
            self.log_message(f"打开报告文件失败: {str(e)}")
    
    def start_bug_analysis(self):
        """开始Bug级别分析"""
        # 查找最新的详细分析报告
        validator = DataValidator()
        latest_file = validator.find_latest_report("详细分析报告")
        
        if not latest_file:
            messagebox.showwarning("警告", "未找到详细分析报告，请先执行开始分析")
            return
        
        # 禁用按钮，启动进度条
        self.bug_analysis_button.config(state='disabled')
        self.progress.start()
        
        # 清空之前的Bug统计结果
        self.clear_bug_stats()
        
        # 在新线程中执行Bug级别分析
        bug_analysis_thread = threading.Thread(target=self.perform_bug_analysis, args=(str(latest_file),))
        bug_analysis_thread.daemon = True
        bug_analysis_thread.start()
    
    def perform_bug_analysis(self, report_file_path):
        """执行Bug级别分析（在后台线程中运行）"""
        try:
            self.log_message("开始Bug级别分析...")
            
            # 执行Bug级别分析
            bug_stats = self.analyze_bug_levels_for_gui(report_file_path)
            
            # 生成Bug级别分析报告
            if bug_stats is not None:
                self.generate_bug_analysis_report(bug_stats, report_file_path)
            
            # 在主线程中更新GUI
            self.root.after(0, lambda: self.update_bug_stats(bug_stats))
            self.root.after(0, lambda: self.bug_analysis_complete("Bug级别分析完成！"))
            
        except Exception as e:
            error_msg = f"Bug级别分析过程中出现错误: {str(e)}"
            self.root.after(0, lambda: self.bug_analysis_complete(error_msg))
    
    def generate_bug_analysis_report(self, bug_stats, source_file_path):
        """生成Bug级别分析报告"""
        try:
            # 确保output目录存在
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            
            # 生成报告文件名
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f"Bug级别分析报告_{timestamp}.xlsx"
            report_path = output_dir / report_filename
            
            # 创建Excel写入器
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                # 写入Bug统计数据
                bug_stats_df = bug_stats.reset_index()
                bug_stats_df.to_excel(writer, sheet_name='Bug级别统计', index=False)
                
                # 获取工作表对象进行格式化
                worksheet = writer.sheets['Bug级别统计']
                
                # 设置列宽
                column_widths = {
                    'A': 25,  # 文件名称
                    'B': 10,  # 总计
                    'C': 12,  # 程序Bug数
                    'D': 15,  # 程序Bug修复数
                    'E': 15,  # 非程序Bug数
                    'F': 18,  # 非程序Bug修复数
                    'G': 8,   # S级
                    'H': 8,   # A级
                    'I': 8,   # B级
                    'J': 8,   # C级
                    'K': 10   # 未分级
                }
                
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                
                # 添加总计行（如果有多个文件）
                if len(bug_stats) > 1:
                    total_row = len(bug_stats) + 2  # +2 因为有标题行和从1开始计数
                    
                    # 计算各列总计
                    worksheet[f'A{total_row}'] = '总计'
                    worksheet[f'B{total_row}'] = bug_stats['总计'].sum()
                    worksheet[f'C{total_row}'] = bug_stats['程序Bug数'].sum() if '程序Bug数' in bug_stats.columns else 0
                    worksheet[f'D{total_row}'] = bug_stats['程序Bug修复数'].sum() if '程序Bug修复数' in bug_stats.columns else 0
                    worksheet[f'E{total_row}'] = bug_stats['非程序Bug数'].sum() if '非程序Bug数' in bug_stats.columns else 0
                    worksheet[f'F{total_row}'] = bug_stats['非程序Bug修复数'].sum() if '非程序Bug修复数' in bug_stats.columns else 0
                    
                    # 计算各级别总计
                    for col_idx, level in enumerate(['S级', 'A级', 'B级', 'C级', '未分级'], start=7):
                        col_letter = chr(ord('A') + col_idx)
                        worksheet[f'{col_letter}{total_row}'] = bug_stats[level].sum() if level in bug_stats.columns else 0
                    
                    # 设置总计行样式
                    from openpyxl.styles import Font, PatternFill
                    bold_font = Font(bold=True)
                    gray_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
                    
                    for col_idx in range(len(bug_stats.columns) + 1):  # +1 for index column
                        col_letter = chr(ord('A') + col_idx)
                        cell = worksheet[f'{col_letter}{total_row}']
                        cell.font = bold_font
                        cell.fill = gray_fill
                
                # 添加分析摘要工作表
                summary_data = []
                summary_data.append(['分析时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                summary_data.append(['源文件', os.path.basename(source_file_path)])
                summary_data.append(['分析文件数量', len(bug_stats)])
                summary_data.append(['总Bug数量', bug_stats['总计'].sum()])
                
                # 各级别统计
                for level in ['S级', 'A级', 'B级', 'C级', '未分级']:
                    if level in bug_stats.columns:
                        count = bug_stats[level].sum()
                        if count > 0:
                            summary_data.append([f'{level}Bug数量', count])
                
                # 程序Bug统计
                if '程序Bug数' in bug_stats.columns:
                    program_bugs = bug_stats['程序Bug数'].sum()
                    program_bugs_fixed = bug_stats['程序Bug修复数'].sum()
                    summary_data.append(['程序Bug总数', program_bugs])
                    summary_data.append(['程序Bug已修复', program_bugs_fixed])
                    if program_bugs > 0:
                        fix_rate = (program_bugs_fixed / program_bugs) * 100
                        summary_data.append(['程序Bug修复率', f'{fix_rate:.1f}%'])
                
                # 非程序Bug统计
                if '非程序Bug数' in bug_stats.columns:
                    non_program_bugs = bug_stats['非程序Bug数'].sum()
                    non_program_bugs_fixed = bug_stats['非程序Bug修复数'].sum()
                    summary_data.append(['非程序Bug总数', non_program_bugs])
                    summary_data.append(['非程序Bug已修复', non_program_bugs_fixed])
                    if non_program_bugs > 0:
                        fix_rate = (non_program_bugs_fixed / non_program_bugs) * 100
                        summary_data.append(['非程序Bug修复率', f'{fix_rate:.1f}%'])
                
                summary_df = pd.DataFrame(summary_data, columns=['项目', '值'])
                summary_df.to_excel(writer, sheet_name='分析摘要', index=False)
                
                # 设置摘要工作表列宽
                summary_worksheet = writer.sheets['分析摘要']
                summary_worksheet.column_dimensions['A'].width = 20
                summary_worksheet.column_dimensions['B'].width = 25
            
            self.log_message(f"Bug级别分析报告已生成: {report_filename}")
            
            # 不自动打开报告，只记录生成信息
            
        except Exception as e:
            self.log_message(f"生成Bug级别分析报告失败: {str(e)}")

    def bug_analysis_complete(self, message):
        """Bug级别分析完成"""
        self.progress.stop()
        self.bug_analysis_button.config(state='normal')
        self.log_message(message)
        
        # 只显示错误提示框，取消完成提示框
        if "错误" in message:
            messagebox.showerror("错误", message)
    
    def analysis_complete(self, message):
        """分析完成"""
        self.progress.stop()
        self.analyze_button.config(state='normal')
        self.log_message(message)
        
        # 只显示错误提示框，取消完成提示框
        if "错误" in message:
            messagebox.showerror("错误", message)

class GUILogHandler(logging.Handler):
    """自定义日志处理器，将日志输出到GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        # 在主线程中更新GUI
        self.text_widget.after(0, lambda: self._append_log(msg))
    
    def _append_log(self, msg):
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)

def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置主题样式
    style = ttk.Style()
    if 'winnative' in style.theme_names():
        style.theme_use('winnative')
    
    # 创建应用
    app = ExcelAnalysisGUI(root)
    
    # 运行应用
    root.mainloop()

if __name__ == "__main__":
    main()