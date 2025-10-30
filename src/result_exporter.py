"""
结果导出模块
支持导出结果到 Excel 文件
"""

import pandas as pd
import logging
from pathlib import Path
from datetime import datetime
from typing import List, Dict

logger = logging.getLogger(__name__)


class ResultExporter:
    """结果导出器"""

    def __init__(self, output_file: str):
        """
        初始化结果导出器

        Args:
            output_file: 输出文件路径（可包含 {timestamp} 占位符）
        """
        # 替换时间戳占位符
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_file = Path(output_file.replace("{timestamp}", timestamp))

        # 确保输出目录存在
        self.output_file.parent.mkdir(parents=True, exist_ok=True)

    def export_results(self, results: List[Dict]) -> bool:
        """
        导出结果到 Excel 文件

        Args:
            results: 结果列表，每个结果是一个字典

        Returns:
            True 如果导出成功
        """
        try:
            if not results:
                logger.warning("结果列表为空，无需导出")
                return False

            # 创建 DataFrame
            df = pd.DataFrame(results)

            # 导出到 Excel
            df.to_excel(self.output_file, index=False, engine='openpyxl')

            logger.info(f"成功导出 {len(results)} 条结果到: {self.output_file}")
            return True

        except Exception as e:
            logger.error(f"导出结果失败: {e}")
            return False

    def append_result(self, result: Dict, all_results: List[Dict]) -> None:
        """
        添加单条结果到结果列表

        Args:
            result: 单条结果字典
            all_results: 所有结果列表
        """
        all_results.append(result)
        logger.debug(f"添加结果: {result}")

    def create_result_entry(
        self,
        row_data: Dict,
        status: str,
        message: str = "",
        **extra_fields
    ) -> Dict:
        """
        创建结果条目

        Args:
            row_data: 原始行数据
            status: 状态（成功/失败）
            message: 消息说明
            **extra_fields: 额外字段

        Returns:
            结果字典
        """
        result = {
            **row_data,  # 包含原始数据
            "处理状态": status,
            "处理消息": message,
            "处理时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            **extra_fields  # 包含额外字段
        }

        return result
