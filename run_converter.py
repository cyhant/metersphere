#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MeterSphere 用例格式转换启动器
双击运行此文件进行转换
"""

import os
import sys
import subprocess


def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    print("\n" + "=" * 50)
    print("  MeterSphere 用例格式转换工具")
    print("=" * 50 + "\n")

    # 检查转换脚本
    converter = os.path.join(script_dir, "convert_metersphere_case.py")
    if not os.path.exists(converter):
        print("[错误] 未找到 convert_metersphere_case.py")
        print("请确保该文件与本脚本在同一目录\n")
        input("按回车键退出...")
        return

    # 查找xlsx文件
    xlsx_files = [
        f
        for f in os.listdir(script_dir)
        if f.endswith(".xlsx") and "_converted" not in f
    ]

    if not xlsx_files:
        print("[错误] 当前目录未找到xlsx文件")
        print("请将要转换的xlsx文件放在本脚本同目录下\n")
        input("按回车键退出...")
        return

    # 选择文件
    if len(xlsx_files) == 1:
        input_file = xlsx_files[0]
        print(f"找到文件: {input_file}")
    else:
        print("找到多个xlsx文件，请选择:\n")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  [{i}] {f}")
        print()

        try:
            choice = int(input("请输入序号: "))
            if 1 <= choice <= len(xlsx_files):
                input_file = xlsx_files[choice - 1]
            else:
                print("\n[错误] 无效选择\n")
                input("按回车键退出...")
                return
        except ValueError:
            print("\n[错误] 请输入数字\n")
            input("按回车键退出...")
            return

    # 执行转换
    print(f"\n正在转换: {input_file}\n")

    try:
        result = subprocess.run(
            [sys.executable, converter, input_file],
            cwd=script_dir,
            capture_output=False,
        )

        if result.returncode == 0:
            output_file = input_file.replace(".xlsx", "_converted.xlsx")
            print(f"\n转换完成! 输出文件: {output_file}\n")

            # 询问是否打开文件
            open_file = input("是否打开输出文件? (y/n): ").strip().lower()
            if open_file == "y":
                output_path = os.path.join(script_dir, output_file)
                os.startfile(output_path)
        else:
            print("\n[错误] 转换失败\n")

    except Exception as e:
        print(f"\n[错误] {e}\n")

    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
