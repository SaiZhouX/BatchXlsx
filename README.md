# Excel批量处理与Bug分析工具

一个功能强大的Excel文件批量处理和Bug数据分析工具，提供直观的图形界面，支持单个文件和批量文件处理，自动生成详细的分析报告。

## 🚀 主要功能

### 1. 智能文件处理
- **单个文件分析**：针对单个Excel文件进行深度分析和数据清理
- **批量文件处理**：同时处理多个Excel文件，自动合并数据
- **智能模式切换**：根据文件数量自动选择最优处理方式

### 2. 数据分析与统计
- **Bug级别分析**：自动识别和统计S级、A级、B级、C级Bug
- **程序Bug统计**：区分程序Bug和非程序Bug，统计修复情况
- **数据质量检查**：自动清理空行、空列和无效数据
- **业务指标统计**：提供全面的数据统计和分析

### 3. 专业报告生成
- **详细分析报告**：包含数据统计、格式化表格和业务分析
- **Bug级别分析报告**：专门的Bug统计和分析报告
- **Excel格式输出**：专业的Excel报告，支持样式和格式化

### 4. 用户友好界面
- **直观的GUI界面**：简单易用的图形界面
- **实时进度显示**：处理过程可视化
- **详细日志记录**：完整的操作日志和错误信息
- **自动报告打开**：处理完成后自动打开生成的报告

## 📁 项目结构

```
BatchXlsx/
├── main.py                    # 主程序（GUI界面）
├── excel_processor.py         # Excel处理基础类
├── batch_processor.py         # 批量文件处理器
├── single_processor.py        # 单个文件处理器
├── bug_analyzer.py            # Bug数据分析器
├── data_validator.py          # 数据验证器
├── config.py                  # 配置文件
├── requirements.txt           # 依赖包列表
├── dist/                      # 打包后的可执行文件
│   └── Excel批量处理与Bug分析工具.exe
└── output/                    # 生成的报告文件
```

## 🛠️ 安装与使用

### 方式一：直接使用exe文件（推荐）
1. 下载 `dist/Excel批量处理与Bug分析工具.exe`
2. 双击运行，无需安装Python环境

### 方式二：从源码运行
1. 克隆项目：
   ```bash
   git clone https://github.com/SaiZhouX/BatchXlsx.git
   cd BatchXlsx
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序：
   ```bash
   python main.py
   ```

## 📖 使用指南

### 1. 添加文件
- **添加单个Excel文件**：点击"添加单个Excel文件"按钮，选择要分析的Excel文件
- **添加文件夹**：点击"添加文件夹"按钮，选择包含Excel文件的文件夹
- **支持多次添加**：可以多次添加文件，构建文件列表

### 2. 开始分析
- 点击"开始分析"按钮开始处理
- 程序会自动判断处理模式：
  - **单个文件**：使用SingleProcessor进行优化处理
  - **多个文件**：使用BatchProcessor进行批量合并处理
- 处理完成后自动打开详细分析报告

### 3. Bug级别分析
- 在完成基础分析后，点击"Bug级别分析"按钮
- 自动分析Bug级别分布和修复情况
- 在界面中查看Bug统计表格
- 生成专门的Bug级别分析报告

### 4. 查看结果
- **Bug级别统计**：在界面表格中查看详细统计
- **处理日志**：在日志标签页查看详细处理信息
- **Excel报告**：自动打开生成的Excel分析报告

## 📊 支持的数据格式

### Excel文件要求
- 支持 `.xlsx` 和 `.xls` 格式
- 包含Bug记录相关的列，如：
  - 编号、严重级别、bug类型
  - 功能模块、修复状态、类型
  - 文件来源信息

### Bug级别格式
- **标准格式**：S-严重、A-重要、B-一般、C-轻微
- **自动转换**：程序会自动转换为S级、A级、B级、C级格式
- **未分级处理**：自动处理空值和未分级数据

### 文件名格式
- 建议包含日期信息（如：0804、20240804）
- 建议包含测试人员姓名
- 程序会自动从文件名提取相关信息

## ⚙️ 配置说明

### 基础配置（config.py）
```python
# 文件夹配置
INPUT_FOLDER = "input"      # 输入文件夹
OUTPUT_FOLDER = "output"    # 输出文件夹

# 支持的文件格式
SUPPORTED_FORMATS = [".xlsx"]

# 报告配置
REPORT_CONFIG = {
    "include_source_column": True,  # 包含文件来源列
    "include_timestamp": True,      # 包含处理时间戳
    "preview_rows": 100,           # 数据预览行数
}
```

## 🔧 技术特性

### 核心技术
- **Python 3.13**：现代Python版本
- **Pandas**：强大的数据处理库
- **OpenPyXL**：Excel文件读写
- **Tkinter**：原生GUI界面

### 性能优化
- **智能模式切换**：根据文件数量选择最优处理方式
- **内存优化**：大文件处理优化
- **多线程处理**：后台处理，界面不卡顿
- **增量处理**：避免重复处理相同文件

### 错误处理
- **完整的异常处理**：捕获和处理各种错误情况
- **详细的错误日志**：帮助定位和解决问题
- **数据验证**：确保数据完整性和一致性

## 📈 生成的报告

### 1. 详细分析报告
- **数据统计**：行数、列数、数据类型统计
- **业务指标**：Bug数量、级别分布、修复率
- **数据预览**：格式化的数据表格
- **质量分析**：空值、重复值分析

### 2. Bug级别分析报告
- **Bug级别统计**：各级别Bug数量统计
- **修复情况分析**：程序Bug和非程序Bug修复率
- **文件来源统计**：按文件来源的Bug分布
- **分析摘要**：关键指标汇总

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进这个项目！

### 开发环境设置
1. Fork项目
2. 创建功能分支：`git checkout -b feature/new-feature`
3. 提交更改：`git commit -am 'Add new feature'`
4. 推送分支：`git push origin feature/new-feature`
5. 创建Pull Request

## 📄 许可证

本项目采用MIT许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 📞 联系方式

如有问题或建议，请通过以下方式联系：
- GitHub Issues: [https://github.com/SaiZhouX/BatchXlsx/issues](https://github.com/SaiZhouX/BatchXlsx/issues)
- 项目主页: [https://github.com/SaiZhouX/BatchXlsx](https://github.com/SaiZhouX/BatchXlsx)

---

**Excel批量处理与Bug分析工具** - 让Excel数据分析更简单、更高效！