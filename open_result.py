"""
快速打开最新的结果文件
"""

import os
from pathlib import Path
import subprocess
import platform


def open_latest_result():
    """打开最新的结果文件"""
    print("=" * 70)
    print("快速打开最新结果文件")
    print("=" * 70)

    # 查找output目录
    output_dir = Path("output")

    if not output_dir.exists():
        print("\n❌ 错误：output目录不存在")
        print(f"   目录路径: {output_dir.absolute()}")
        print("\n可能原因：")
        print("  • 程序还没有执行过")
        print("  • 当前不在项目根目录")
        print("\n解决方法：")
        print("  1. 先运行 python main.py 执行程序")
        print("  2. 或者 cd 到项目根目录")
        return

    # 查找所有结果文件
    result_files = []
    for pattern in ["outpatient_results_*.xlsx", "emergency_results_*.xlsx"]:
        result_files.extend(output_dir.glob(pattern))

    if not result_files:
        print("\n❌ 错误：没有找到结果文件")
        print(f"   搜索目录: {output_dir.absolute()}")
        print("\n可能原因：")
        print("  • 程序还没有执行过")
        print("  • 结果文件被手动删除")
        print("\n解决方法：")
        print("  1. 运行 python main.py 执行程序")
        print("  2. 检查output目录是否为空")
        return

    # 按修改时间排序，获取最新的文件
    latest_file = max(result_files, key=lambda f: f.stat().st_mtime)

    print(f"\n✓ 找到最新的结果文件:")
    print(f"  文件名: {latest_file.name}")
    print(f"  路径:   {latest_file.absolute()}")
    print(f"  大小:   {latest_file.stat().st_size / 1024:.2f} KB")

    # 获取文件修改时间
    import datetime
    mtime = datetime.datetime.fromtimestamp(latest_file.stat().st_mtime)
    print(f"  时间:   {mtime.strftime('%Y-%m-%d %H:%M:%S')}")

    print("\n正在打开...")

    # 打开文件
    try:
        if platform.system() == "Windows":
            # Windows: 打开文件夹并选中文件
            subprocess.run(['explorer', '/select,', str(latest_file)])
            print("✓ 文件夹已打开，文件已选中")
        elif platform.system() == "Darwin":
            # macOS: 打开并选中文件
            subprocess.run(['open', '-R', str(latest_file)])
            print("✓ 文件夹已打开，文件已选中")
        else:
            # Linux: 只能打开文件夹
            subprocess.run(['xdg-open', str(latest_file.parent)])
            print("✓ 文件夹已打开")
            print(f"  请在文件夹中找到: {latest_file.name}")

    except Exception as e:
        print(f"\n❌ 无法自动打开文件夹: {e}")
        print(f"\n请手动打开:")
        print(f"  {latest_file.absolute()}")

    print("\n" + "=" * 70)
    print("提示：")
    print("  • 双击Excel文件查看详细结果")
    print("  • 查看'处理状态'列了解成功/失败情况")
    print("  • 查看'处理消息'列了解失败原因")
    print("=" * 70)


def list_all_results():
    """列出所有结果文件"""
    print("\n" + "=" * 70)
    print("所有结果文件列表")
    print("=" * 70)

    output_dir = Path("output")

    if not output_dir.exists():
        print("\n❌ output目录不存在")
        return

    # 查找所有结果文件
    result_files = []
    for pattern in ["outpatient_results_*.xlsx", "emergency_results_*.xlsx"]:
        result_files.extend(output_dir.glob(pattern))

    if not result_files:
        print("\n❌ 没有找到任何结果文件")
        return

    # 按修改时间排序（最新的在前）
    result_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)

    print(f"\n共找到 {len(result_files)} 个结果文件:\n")

    import datetime
    for i, file in enumerate(result_files, 1):
        mtime = datetime.datetime.fromtimestamp(file.stat().st_mtime)
        size_kb = file.stat().st_size / 1024

        marker = "← 最新" if i == 1 else ""

        print(f"{i}. {file.name} {marker}")
        print(f"   时间: {mtime.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"   大小: {size_kb:.2f} KB")
        print(f"   路径: {file.absolute()}")
        print()

    print("=" * 70)


def main():
    """主函数"""
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "--list":
        # 列出所有结果文件
        list_all_results()
    else:
        # 打开最新的结果文件
        open_latest_result()

    print("\n用法:")
    print("  python open_result.py         # 打开最新结果文件")
    print("  python open_result.py --list  # 列出所有结果文件")
    print()


if __name__ == "__main__":
    main()
