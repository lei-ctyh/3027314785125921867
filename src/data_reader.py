"""
数据读取模块
支持读取 Excel 和 CSV 文件
"""

import pandas as pd
import logging
from pathlib import Path
from typing import List, Dict

logger = logging.getLogger(__name__)


class DataReader:
    """数据读取器"""

    def __init__(self, file_path: str, sheet_name: str = "Sheet1"):
        """
        初始化数据读取器

        Args:
            file_path: 数据文件路径
            sheet_name: Excel sheet 名称（仅用于 Excel 文件）
        """
        self.file_path = Path(file_path)
        self.sheet_name = sheet_name

        if not self.file_path.exists():
            raise FileNotFoundError(f"数据文件不存在: {self.file_path}")

    def read_data(self) -> List[Dict]:
        """
        读取数据文件

        Returns:
            数据列表，每行数据转换为字典

        Raises:
            ValueError: 不支持的文件格式
        """
        file_extension = self.file_path.suffix.lower()

        try:
            if file_extension in ['.xlsx', '.xls']:
                logger.info(f"读取 Excel 文件: {self.file_path}")
                df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
            elif file_extension == '.csv':
                logger.info(f"读取 CSV 文件: {self.file_path}")
                df = pd.read_csv(self.file_path)
            else:
                raise ValueError(f"不支持的文件格式: {file_extension}")

            # 将 DataFrame 转换为字典列表
            data_list = df.to_dict('records')
            logger.info(f"成功读取 {len(data_list)} 条数据")

            return data_list

        except Exception as e:
            logger.error(f"读取数据文件失败: {e}")
            raise

    def validate_columns(self, required_columns: List[str]) -> bool:
        """
        验证数据文件是否包含必需的列

        Args:
            required_columns: 必需的列名列表

        Returns:
            True 如果所有必需列都存在
        """
        try:
            if self.file_path.suffix.lower() in ['.xlsx', '.xls']:
                df = pd.read_excel(self.file_path, sheet_name=self.sheet_name, nrows=1)
            else:
                df = pd.read_csv(self.file_path, nrows=1)

            existing_columns = set(df.columns)
            missing_columns = set(required_columns) - existing_columns

            if missing_columns:
                logger.error(f"数据文件缺少以下列: {missing_columns}")
                return False

            logger.info("数据文件列验证通过")
            return True

        except Exception as e:
            logger.error(f"验证列失败: {e}")
            return False
