"""
表单填写核心模块
"""

import logging
import time
from typing import Dict, Any
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

logger = logging.getLogger(__name__)


class FormFiller:
    """表单填写器"""

    # 定位方式映射
    LOCATOR_MAP = {
        "id": By.ID,
        "name": By.NAME,
        "xpath": By.XPATH,
        "css_selector": By.CSS_SELECTOR,
        "class_name": By.CLASS_NAME,
        "tag_name": By.TAG_NAME,
        "link_text": By.LINK_TEXT,
    }

    def __init__(self, driver, form_elements: Dict, timeout: int = 30):
        """
        初始化表单填写器

        Args:
            driver: WebDriver 实例
            form_elements: 表单元素配置
            timeout: 超时时间（秒）
        """
        self.driver = driver
        self.form_elements = form_elements
        self.timeout = timeout
        self.wait = WebDriverWait(driver, timeout)

    def fill_form(self, data: Dict[str, Any]) -> bool:
        """
        填写表单

        Args:
            data: 表单数据字典，key 为表单字段名，value 为要填写的值

        Returns:
            True 如果填写成功，False 否则
        """
        try:
            logger.info(f"开始填写表单，数据: {data}")

            for field_name, field_value in data.items():
                # 跳过空值
                if field_value is None or str(field_value).strip() == "":
                    logger.debug(f"跳过空字段: {field_name}")
                    continue

                # 获取元素配置
                element_config = self.form_elements.get(field_name)
                if not element_config:
                    logger.warning(f"配置中未找到字段: {field_name}")
                    continue

                # 填写字段
                self._fill_field(field_name, field_value, element_config)
                time.sleep(0.5)  # 短暂延迟，模拟人工操作

            logger.info("表单填写完成")
            return True

        except Exception as e:
            logger.error(f"填写表单失败: {e}")
            return False

    def _fill_field(self, field_name: str, field_value: Any, config: Dict) -> None:
        """
        填写单个字段

        Args:
            field_name: 字段名
            field_value: 字段值
            config: 元素配置
        """
        element_type = config.get("type", "input")
        locator_type = config.get("locator", "id")
        locator_value = config.get("value")

        # 获取定位方式
        by = self.LOCATOR_MAP.get(locator_type, By.ID)

        try:
            # 等待元素可见
            element = self.wait.until(
                EC.presence_of_element_located((by, locator_value))
            )

            # 根据元素类型填写
            if element_type == "input" or element_type == "textarea":
                element.clear()
                element.send_keys(str(field_value))
                logger.debug(f"填写文本字段 {field_name}: {field_value}")

            elif element_type == "select":
                select = Select(element)
                # 尝试按值选择，如果失败则按文本选择
                try:
                    select.select_by_value(str(field_value))
                except:
                    select.select_by_visible_text(str(field_value))
                logger.debug(f"选择下拉框 {field_name}: {field_value}")

            elif element_type == "button":
                element.click()
                logger.debug(f"点击按钮 {field_name}")

            else:
                logger.warning(f"未知元素类型: {element_type}")

        except TimeoutException:
            logger.error(f"超时：找不到元素 {field_name} (定位器: {locator_type}={locator_value})")
            raise
        except Exception as e:
            logger.error(f"填写字段 {field_name} 失败: {e}")
            raise

    def submit_form(self, submit_button_name: str = "submit_button") -> bool:
        """
        提交表单

        Args:
            submit_button_name: 提交按钮在配置中的名称

        Returns:
            True 如果提交成功
        """
        try:
            button_config = self.form_elements.get(submit_button_name)
            if not button_config:
                logger.error(f"配置中未找到提交按钮: {submit_button_name}")
                return False

            locator_type = button_config.get("locator", "id")
            locator_value = button_config.get("value")
            by = self.LOCATOR_MAP.get(locator_type, By.ID)

            # 等待按钮可点击
            button = self.wait.until(
                EC.element_to_be_clickable((by, locator_value))
            )

            button.click()
            logger.info("表单提交成功")
            time.sleep(2)  # 等待提交完成
            return True

        except Exception as e:
            logger.error(f"提交表单失败: {e}")
            return False

    def check_success(self, success_indicator: Dict) -> bool:
        """
        检查提交是否成功

        Args:
            success_indicator: 成功指示器配置，如 {"type": "url_contains", "value": "success"}
                            或 {"type": "element_text", "locator": "id", "value": "success_msg"}

        Returns:
            True 如果检测到成功标志
        """
        try:
            indicator_type = success_indicator.get("type")

            if indicator_type == "url_contains":
                expected_url = success_indicator.get("value")
                current_url = self.driver.current_url
                return expected_url in current_url

            elif indicator_type == "element_text":
                locator_type = success_indicator.get("locator", "id")
                locator_value = success_indicator.get("value")
                by = self.LOCATOR_MAP.get(locator_type, By.ID)

                element = self.wait.until(
                    EC.presence_of_element_located((by, locator_value))
                )
                return element.text != ""

            else:
                logger.warning(f"未知的成功指示器类型: {indicator_type}")
                return False

        except Exception as e:
            logger.error(f"检查成功状态失败: {e}")
            return False
