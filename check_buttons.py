"""
快速验证GUI按钮是否可见
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

import tkinter as tk

def check_button_visibility():
    """检查按钮可见性"""
    print("=" * 70)
    print("GUI按钮可见性测试")
    print("=" * 70)
    print("\n即将打开配置窗口...")
    print("\n请检查以下事项：")
    print("✓ 窗口右侧是否有垂直滚动条")
    print("✓ 滚动到最底部，是否能看到两个按钮")
    print("  - 右侧：绿色的'开始执行'按钮")
    print("  - 左侧：灰色的'取消'按钮")
    print("✓ 可以使用鼠标滚轮上下滚动")
    print("✓ 可以拖动窗口边缘调整大小")
    print("\n如果看不到按钮，请滚动到窗口最底部！")
    print("=" * 70)

    input("\n按回车键打开窗口...")

    from gui import show_config_gui

    config = show_config_gui()

    if config:
        print("\n" + "=" * 70)
        print("✓ 成功！用户点击了'开始执行'按钮")
        print("=" * 70)
        print("\n收到的配置:")
        for key, value in config.items():
            if key == 'password':
                print(f"  {key}: {'*' * len(value)}")
            else:
                print(f"  {key}: {value}")
    else:
        print("\n" + "=" * 70)
        print("✓ 用户点击了'取消'按钮或关闭了窗口")
        print("=" * 70)

if __name__ == "__main__":
    check_button_visibility()
