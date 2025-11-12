import os
import shutil
from PyInstaller.__main__ import run

def clean_old_builds():
    """清理旧的build和dist目录"""
    for dir_name in ["build", "dist"]:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已清理目录: {dir_name}")

def build_exe():
    """调用PyInstaller API打包程序"""
    # 打包参数配置
    params = [
        "main.py",  # 程序入口文件
        "--name=WordProcessor",  # 生成的可执行文件名称
        "--icon=gh.ico",  # 程序图标（使用项目根目录的gh.ico）
        "--onefile",  # 打包为单个可执行文件
        "--windowed",  # 无控制台窗口（GUI程序推荐）
        # 如需排除不必要的库，可加--exclude-module参数，例如：
        # "--exclude-module=some_unused_module",
    ]
    run(params)
    print("打包完成！可执行文件位于dist目录")

if __name__ == "__main__":
    clean_old_builds()
    build_exe()