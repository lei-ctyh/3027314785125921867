"""
快速测试GUI按钮显示
验证所有按钮是否正确显示和工作
"""

import sys
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

import tkinter as tk
from gui import ConfigGUI


def test_buttons():
    """测试按钮显示"""
    print("=" * 60)
    print("GUI按钮测试")
    print("=" * 60)
    print("\n请检查配置界面中的按钮：")
    print("1. 底部应该有两个按钮")
    print("2. 右侧是绿色的'开始执行'按钮")
    print("3. 左侧是灰色的'取消'按钮")
    print("4. 两个按钮都应该比较大且醒目")
    print("5. 鼠标悬停时光标应该变成手型")
    print("\n窗口尺寸: 650x750")
    print("=" * 60)
    print("\n打开GUI窗口...")

    gui = ConfigGUI()

    # 添加一些调试信息
    root = gui.root
    print(f"窗口大小: {root.geometry()}")

    config = gui.show()

    if config:
        print("\n✓ 用户点击了'开始执行'按钮")
        print(f"✓ 收到配置: {config}")
    else:
        print("\n✓ 用户点击了'取消'按钮")


if __name__ == "__main__":
    test_buttons()
