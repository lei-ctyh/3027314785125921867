"""
登录处理模块
"""

import logging
import time
from typing import Dict
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

logger = logging.getLogger(__name__)


class LoginHandler:
    """登录处理器"""

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

    def __init__(self, driver, login_config: Dict, timeout: int = 30):
        """
        初始化登录处理器

        Args:
            driver: WebDriver 实例
            login_config: 登录配置
            timeout: 超时时间（秒）
        """
        self.driver = driver
        self.login_config = login_config
        self.timeout = timeout
        self.wait = WebDriverWait(driver, timeout)

    def login(self) -> bool:
        """
        执行登录

        Returns:
            True 如果登录成功，False 否则
        """
        try:
            logger.info("开始执行登录流程...")

            # 打开登录页面
            login_url = self.login_config.get("login_url")
            logger.info(f"打开登录页面: {login_url}")
            self.driver.get(login_url)
            time.sleep(2)  # 等待页面加载

            # 获取登录信息
            username = self.login_config.get("username")
            password = self.login_config.get("password")
            elements_config = self.login_config.get("elements", {})

            if not username or not password:
                logger.error("登录配置中缺少用户名或密码")
                return False

            # 填写用户名
            username_config = elements_config.get("username_field", {})
            self._fill_field(username, username_config, "用户名")

            time.sleep(0.5)

            # 填写密码
            password_config = elements_config.get("password_field", {})
            self._fill_field(password, password_config, "密码")

            time.sleep(0.5)

            # 点击登录按钮
            login_button_config = elements_config.get("login_button", {})
            self._click_button(login_button_config, "登录按钮")

            logger.info("等待登录响应...")
            time.sleep(3)  # 等待登录处理

            # 验证登录是否成功
            success = self._verify_login_success()

            if success:
                logger.info("登录成功！")
                return True
            else:
                logger.error("登录失败")
                return False

        except Exception as e:
            logger.error(f"登录过程发生错误: {e}")
            return False

    def _fill_field(self, value: str, config: Dict, field_name: str) -> None:
        """
        填写输入框

        Args:
            value: 要填写的值
            config: 元素配置
            field_name: 字段名称（用于日志）
        """
        locator_type = config.get("locator", "id")
        locator_value = config.get("value")

        by = self.LOCATOR_MAP.get(locator_type, By.ID)

        try:
            element = self.wait.until(
                EC.presence_of_element_located((by, locator_value))
            )
            element.clear()
            element.send_keys(value)
            logger.debug(f"填写{field_name}: {value}")

        except TimeoutException:
            logger.error(f"超时：找不到{field_name}输入框 (定位器: {locator_type}={locator_value})")
            raise
        except Exception as e:
            logger.error(f"填写{field_name}失败: {e}")
            raise

    def _click_button(self, config: Dict, button_name: str) -> None:
        """
        点击按钮

        Args:
            config: 按钮配置
            button_name: 按钮名称（用于日志）
        """
        locator_type = config.get("locator", "id")
        locator_value = config.get("value")

        by = self.LOCATOR_MAP.get(locator_type, By.ID)

        try:
            button = self.wait.until(
                EC.element_to_be_clickable((by, locator_value))
            )
            button.click()
            logger.debug(f"点击{button_name}")

        except TimeoutException:
            logger.error(f"超时：找不到{button_name} (定位器: {locator_type}={locator_value})")
            raise
        except Exception as e:
            logger.error(f"点击{button_name}失败: {e}")
            raise

    def _verify_login_success(self) -> bool:
        """
        验证登录是否成功

        Returns:
            True 如果验证成功
        """
        try:
            success_indicator = self.login_config.get("success_indicator", {})
            indicator_type = success_indicator.get("type")
            indicator_value = success_indicator.get("value")

            if indicator_type == "url_contains":
                # 检查URL是否包含特定字符串
                current_url = self.driver.current_url
                logger.debug(f"当前URL: {current_url}")

                if indicator_value in current_url:
                    logger.info(f"URL验证成功: URL包含 '{indicator_value}'")
                    return True
                else:
                    logger.warning(f"URL验证失败: URL不包含 '{indicator_value}'")
                    return False

            elif indicator_type == "element_exists":
                # 检查页面是否存在特定元素
                locator_type = success_indicator.get("locator", "id")
                by = self.LOCATOR_MAP.get(locator_type, By.ID)

                try:
                    self.wait.until(
                        EC.presence_of_element_located((by, indicator_value))
                    )
                    logger.info(f"元素验证成功: 找到元素 '{indicator_value}'")
                    return True
                except TimeoutException:
                    logger.warning(f"元素验证失败: 找不到元素 '{indicator_value}'")
                    return False

            else:
                logger.warning(f"未知的验证类型: {indicator_type}，默认认为登录成功")
                return True

        except Exception as e:
            logger.error(f"验证登录状态失败: {e}")
            return False
