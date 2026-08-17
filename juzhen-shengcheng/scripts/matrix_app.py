# -*- coding: utf-8 -*-
"""矩阵生成工具 EXE 入口 (PyInstaller 打包用)
用法: 矩阵生成工具.exe 打标表.xlsx   (或双击后拖入文件)
"""
import os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import run_analysis


def main():
    # 控制台编码保险 (GBK 控制台无法输出部分 Unicode, 避免崩溃)
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    if len(sys.argv) > 1:
        path = sys.argv[1].strip().strip('"')
    else:
        try:
            path = input('请把【打标后的Excel文件】拖进来后回车: ').strip().strip('"')
        except EOFError:
            print('未提供文件')
            return
    if not path or not os.path.exists(path):
        print('[X] 文件不存在:', path)
        return
    sys.argv = ['app', '--input', path]
    run_analysis.main()
    print()
    print('============================================')
    print('  完成! 输出文件与打标表同目录')
    print('============================================')
    try:
        input('按回车退出...')
    except EOFError:
        pass


if __name__ == '__main__':
    main()
