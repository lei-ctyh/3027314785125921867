"""
测试GUI模块
仅测试GUI界面的显示和输入收集功能
"""

import sys
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from gui import show_config_gui, show_confirmation_dialog


def test_config_gui():
    """测试配置GUI"""
    print("=" * 60)
    print("测试配置输入界面")
    print("=" * 60)

    config = show_config_gui()

    if config:
        print("\n收到用户配置:")
        print(f"  用户名: {config['username']}")
        print(f"  密码: {'*' * len(config['password'])}")
        print(f"  月份: {config['month']}")
        print(f"  功能类型: {config['function_type']}")
        print(f"  输入文件: {config['input_file']}")
        print(f"  无头模式: {config['headless']}")
        print("\n✓ 配置GUI测试通过")
        return True
    else:
        print("\n✗ 用户取消操作")
        return False


def test_confirmation_dialog():
    """测试确认对话框"""
    print("\n" + "=" * 60)
    print("测试确认对话框")
    print("=" * 60)

    message = """请在浏览器中确认以下信息：

1. 请手动输入报表日期
2. 请手动输入门诊总量
3. 确认所有信息无误后，点击"我已完成，继续"按钮

程序将继续自动填写表单数据。"""

    confirmed = show_confirmation_dialog("测试确认", message)

    if confirmed:
        print("\n✓ 用户确认")
        return True
    else:
        print("\n✗ 用户未确认")
        return False


def main():
    """主测试函数"""
    print("\n开始GUI模块测试\n")

    # 测试配置GUI
    if not test_config_gui():
        print("\n测试中断")
        return

    # 测试确认对话框
    test_confirmation_dialog()

    print("\n" + "=" * 60)
    print("所有测试完成")
    print("=" * 60)


if __name__ == "__main__":
    main()
