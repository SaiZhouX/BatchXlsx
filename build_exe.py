"""
Excel批量处理工具打包脚本
使用PyInstaller将Python项目打包为exe文件
"""

import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

def build_exe():
    """构建exe文件"""
    
    # 确保在正确的目录
    project_dir = Path(__file__).parent
    os.chdir(project_dir)
    
    print("开始打包Excel批量处理工具...")
    
    # 生成带时间戳的文件名
    timestamp = datetime.now().strftime('%Y%m%d')
    exe_name = f"BatchXlsxTool_{timestamp}"
    print(f"生成的exe文件名: {exe_name}.exe")
    
    # PyInstaller命令参数
    cmd = [
        'pyinstaller',
        '--onefile',                    # 打包成单个exe文件
        '--windowed',                   # 不显示控制台窗口
        f'--name={exe_name}',           # exe文件名（带时间戳）
        '--icon=icon.ico',              # 图标文件（如果存在）
        '--add-data=config;config',     # 添加配置文件夹
        '--hidden-import=openpyxl',     # 隐式导入
        '--hidden-import=pandas',
        '--hidden-import=tkinter',
        '--hidden-import=xlrd',
        '--hidden-import=xlsxwriter',
        '--clean',                      # 清理临时文件
        'main.py'                       # 主程序文件
    ]
    
    # 如果没有图标文件，移除图标参数
    if not Path('icon.ico').exists():
        cmd.remove('--icon=icon.ico')
    
    # 如果没有config文件夹，移除配置参数
    if not Path('config').exists():
        cmd.remove('--add-data=config;config')
    
    try:
        # 执行打包命令
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("打包成功！")
        print(f"exe文件位置: {project_dir}/dist/{exe_name}.exe")
        
        # 显示文件大小
        exe_path = project_dir / "dist" / f"{exe_name}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"文件大小: {size_mb:.1f} MB")
        
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False
    
    return exe_name

def clean_build_files():
    """清理构建文件"""
    import shutil
    
    dirs_to_remove = ['build', '__pycache__']
    
    # 动态查找所有带时间戳的spec文件
    spec_files = list(Path('.').glob('BatchXlsxTool_*.spec'))
    
    for dir_name in dirs_to_remove:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"已删除: {dir_name}")
    
    for spec_file in spec_files:
        if spec_file.exists():
            spec_file.unlink()
            print(f"已删除: {spec_file.name}")

if __name__ == "__main__":
    print("=" * 50)
    print("Excel批量处理工具 - 打包脚本")
    print("=" * 50)
    
    # 检查依赖
    try:
        import PyInstaller
        print(f"PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("错误: 未安装PyInstaller")
        print("请运行: pip install pyinstaller")
        sys.exit(1)
    
    # 执行打包
    result = build_exe()
    
    if result:
        print("\n" + "=" * 50)
        print("打包完成！")
        print(f"exe文件位置: dist/{result}.exe")
        print("=" * 50)
        
        # 自动清理构建文件
        print("\n正在清理构建文件...")
        clean_build_files()
        print("构建文件已清理")
    else:
        print("\n打包失败，请检查错误信息")