#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成示例 Excel 数据文件
"""

import pandas as pd
from pathlib import Path

def create_sample_data():
    """创建示例数据文件"""
    # 示例数据
    data = {
        'name': ['张三', '李四', '王五'],
        'email': ['zhangsan@example.com', 'lisi@example.com', 'wangwu@example.com'],
        'phone': ['13800138000', '13900139000', '13700137000'],
        'gender': ['男', '女', '男'],
        'message': ['测试消息1', '测试消息2', '测试消息3']
    }

    # 创建 DataFrame
    df = pd.DataFrame(data)

    # 确保 data 目录存在
    data_dir = Path(__file__).parent / "data"
    data_dir.mkdir(exist_ok=True)

    # 保存为 Excel 文件
    output_file = data_dir / "input_data.xlsx"
    df.to_excel(output_file, index=False, engine='openpyxl')

    print(f"示例数据文件创建成功: {output_file}")
    print(f"\n数据预览:")
    print(df.to_string(index=False))
    print(f"\n共 {len(df)} 行数据")

if __name__ == "__main__":
    try:
        create_sample_data()
    except Exception as e:
        print(f"创建失败: {e}")
