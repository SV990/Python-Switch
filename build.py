import PyInstaller.__main__
import os
from pathlib import Path

# 获取当前目录
current_dir = Path(__file__).parent

# 配置打包参数
params = [
    'switch.py',  # 主程序文件
    '--name=Python版本切换器',  # 程序名称
    '--noconsole',  # 不显示控制台
    '--onefile',  # 打包成单个文件
    f'--icon={current_dir / "python.ico"}',  # 图标
    '--uac-admin',  # 请求管理员权限
    '--add-data=python.ico;.',  # 添加图标文件
    '--clean',  # 清理临时文件
    '--version-file=version.txt',  # 版本信息文件
]

# 执行打包
PyInstaller.__main__.run(params)
