"""
测试确认对话框显示
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from gui import show_confirmation_dialog


def test_first_confirmation():
    """测试开始前确认对话框"""
    print("=" * 70)
    print("测试：开始前确认对话框")
    print("=" * 70)

    message = """请在浏览器中确认以下信息：

1. 请手动输入报表日期
2. 请手动输入门诊总量
3. 确认所有信息无误后，点击"我已完成，继续"按钮

程序将继续自动填写表单数据。"""

    print("\n即将显示对话框...")
    print("\n请检查：")
    print("✓ 所有文字是否完整显示")
    print("✓ 窗口右侧是否有滚动条")
    print("✓ 底部是否有蓝色的'我已完成，继续'按钮")
    print("✓ 可以使用鼠标滚轮滚动查看所有内容")
    print("✓ 可以拖动窗口边缘调整大小")
    print("-" * 70)

    input("按回车键显示对话框...")

    confirmed = show_confirmation_dialog("开始前确认", message)

    if confirmed:
        print("\n✓ 用户点击了'我已完成，继续'按钮")
    else:
        print("\n✗ 对话框被关闭")

    return confirmed


def test_final_confirmation():
    """测试完成后确认对话框"""
    print("\n" + "=" * 70)
    print("测试：完成后确认对话框")
    print("=" * 70)

    message = """数据填写完成！

总计：100 条
成功：95 条
失败：5 条

请在浏览器中检查填写结果，确认无误后：
1. 手动点击"上报"按钮
2. 等待上报完成
3. 点击"我已完成，继续"按钮关闭程序"""

    print("\n即将显示对话框...")
    print("\n请检查：")
    print("✓ 所有统计信息是否完整显示")
    print("✓ 所有操作步骤是否清晰可见")
    print("✓ 底部按钮是否可见")
    print("-" * 70)

    input("按回车键显示对话框...")

    confirmed = show_confirmation_dialog("上报确认", message)

    if confirmed:
        print("\n✓ 用户点击了'我已完成，继续'按钮")
    else:
        print("\n✗ 对话框被关闭")

    return confirmed


def main():
    """主测试函数"""
    print("\n确认对话框显示测试\n")

    # 测试第一个确认对话框
    if not test_first_confirmation():
        print("\n测试中断")
        return

    # 测试第二个确认对话框
    test_final_confirmation()

    print("\n" + "=" * 70)
    print("所有测试完成")
    print("=" * 70)
    print("\n如果内容显示不全，请：")
    print("1. 使用鼠标滚轮向下滚动")
    print("2. 拖动右侧的滚动条")
    print("3. 拖动窗口边缘放大窗口")
    print("=" * 70)


if __name__ == "__main__":
    main()
