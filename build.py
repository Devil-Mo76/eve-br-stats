# -*- coding: utf-8 -*-
"""
PyInstaller 打包脚本
--------------------------------------------------------------------------------
用法：
    pip install pyinstaller
    python build.py

输出：
    dist/BRStatisticsTool.exe
--------------------------------------------------------------------------------
通用版改动说明：
    1. EXE 名称、图标、应用名统一从 config.py / 本文件顶部常量读取，不再散落硬编码；
    2. 图标默认使用 app_icon.ico（可用 tools/make_icon.py 重新生成）。
"""

import os
import shutil
import sys

import PyInstaller.__main__

import config as cfg

# ============================ 需要你确认的部分 ============================
# 【需修改】EXE 文件名称
EXE_NAME = "BRStatisticsTool"

# 【需修改】EXE 图标（必须是 .ico 格式，推荐 256x256）
#          想换成自己的图标：把 .ico 放到项目根目录，然后改这里的文件名
ICON_FILE = "app_icon.ico"

# 【需修改】Windows 文件属性里显示的应用名称
APP_NAME = cfg.SITE_NAME

# 【需修改】EXE 版本信息文件（右键 EXE -> 属性 -> 详细信息）
VERSION_FILE = "version_info.txt"
# =========================================================================

# 打包时一并带上的资源目录
DATA_FOLDERS = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('data', 'data'),
]

# 需要显式声明的隐藏导入（PyInstaller 有时扫不到）
HIDDEN_IMPORTS = [
    'jinja2',
    'werkzeug',
    'flask',
    'openpyxl',
    'openpyxl.chart.label',
    'openpyxl.chart.legend',
    'pandas',
]


def build_app():
    try:
        print(f"=== 正在构建 {APP_NAME} EXE 版本 ===")

        if not os.path.exists(ICON_FILE):
            print(f"警告: 未找到图标文件 {ICON_FILE}，将使用默认图标")

        # 1. 清理上次构建产物
        print("清理之前的构建文件...")
        for path in ('build', 'dist'):
            if os.path.exists(path):
                shutil.rmtree(path)
        spec_file = f'{EXE_NAME}.spec'
        if os.path.exists(spec_file):
            os.remove(spec_file)

        # 2. 组装 PyInstaller 参数
        args = [
            'app.py',
            '--onefile',
            '--console',
            f'--name={EXE_NAME}',
            '--clean',
            '--noconfirm',
            '--workpath=build',
            '--distpath=dist',
            '--log-level=WARN',
            f'--icon={ICON_FILE}',
        ]

        for src, dst in DATA_FOLDERS:
            args.append('--add-data')
            # Windows 用 ; 分隔，其它平台用 :
            args.append(f'{src};{dst}' if sys.platform == 'win32' else f'{src}:{dst}')

        for imp in HIDDEN_IMPORTS:
            args.append('--hidden-import')
            args.append(imp)

        # 【需修改】版本信息文件：控制 EXE 属性里的公司名/说明/版权
        if os.path.exists(VERSION_FILE):
            args.append(f'--version-file={VERSION_FILE}')
        else:
            print(f"提示: 未找到 {VERSION_FILE}，EXE 将不带版本信息")

        # 3. 执行打包
        print("开始打包...")
        PyInstaller.__main__.run(args)

        # 4. 结果检查
        print("\n=== 打包完成 ===")
        exe_path = os.path.join('dist', f'{EXE_NAME}.exe')
        if os.path.exists(exe_path):
            print(f"成功生成可执行文件: {exe_path}")
            print(f"文件大小: {os.path.getsize(exe_path) / 1024 / 1024:.2f} MB")
            if os.path.exists(ICON_FILE):
                print(f"已应用自定义图标: {ICON_FILE}")
        else:
            print("构建失败，未生成可执行文件")

    except Exception as e:
        print(f"\n!!! 构建过程中发生错误: {e} !!!")
        sys.exit(1)


if __name__ == '__main__':
    build_app()
