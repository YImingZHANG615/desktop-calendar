# -*- coding: utf-8 -*-
import os
import sys
import subprocess

def build_exe():
    """
    使用PyInstaller将程序打包成EXE文件
    """
    print("开始构建EXE文件...")
    
    # 检查并安装必要的依赖
    dependencies = ["pyinstaller", "pywin32"]
    for dep in dependencies:
        try:
            __import__(dep.replace("-", "_"))
            print(f"已检测到{dep}")
        except ImportError:
            print(f"未检测到{dep}，正在安装...")
            subprocess.call([sys.executable, "-m", "pip", "install", dep])
            print(f"{dep}安装完成")
    
    # 确保数据文件存在
    data_files = ["events_data.json", "event_types.json"]
    for data_file in data_files:
        if not os.path.exists(data_file):
            print(f"创建空的{data_file}文件...")
            with open(data_file, "w", encoding="utf-8") as f:
                f.write("{}")
        else:
            print(f"数据文件{data_file}已存在")
    
    # 构建命令
    cmd = [
        "pyinstaller",
        "--name=桌面日历",
        "--windowed",  # 无控制台窗口
        "--onefile",   # 打包成单个文件
        "--icon=calendar.ico",  # 使用日历图标
        "--add-data=calendar.ico;.",  # 添加图标文件
        "--add-data=events_data.json;.",  # 添加事件数据文件
        "--add-data=event_types.json;.",  # 添加事件类型数据文件
        "--hidden-import=win32com.client",  # 添加隐藏导入
        "--hidden-import=win32api",
        "--hidden-import=win32con",
        "app.py"
    ]
    
    # 执行打包命令
    print("正在执行打包命令...")
    result = subprocess.call(cmd)
    
    if result == 0:
        print("EXE文件构建成功！")
        print("可执行文件位于: dist/桌面日历.exe")
        print("\n特性:")
        print("1. 始终置于底部，不会覆盖其他窗口")
        print("2. 默认开机自启动")
        print("3. 最小化到系统托盘，不占用任务栏")
    else:
        print("构建失败，请检查错误信息")

if __name__ == "__main__":
    build_exe()
