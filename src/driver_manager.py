"""
浏览器驱动管理模块
"""

import logging
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

logger = logging.getLogger(__name__)


class DriverManager:
    """浏览器驱动管理器"""

    def __init__(self, headless: bool = False, window_size: str = "1920,1080"):
        """
        初始化驱动管理器

        Args:
            headless: 是否无头模式
            window_size: 浏览器窗口大小
        """
        self.headless = headless
        self.window_size = window_size
        self.driver = None

    def create_driver(self) -> webdriver.Chrome:
        """
        创建并配置 Chrome 驱动

        Returns:
            配置好的 Chrome WebDriver 实例

        Raises:
            FileNotFoundError: 找不到 chromedriver.exe
            Exception: 启动驱动失败
        """
        # 查找 chromedriver.exe
        project_root = Path(__file__).parent.parent
        driver_path = project_root / "chromedriver-win64" / "chromedriver.exe"

        if not driver_path.exists():
            raise FileNotFoundError(
                f"找不到 ChromeDriver: {driver_path}\n"
                f"请将 chromedriver.exe 放到项目目录"
            )

        logger.info(f"使用 ChromeDriver: {driver_path}")

        # 配置 Chrome 选项
        chrome_options = Options()

        if self.headless:
            chrome_options.add_argument("--headless")
            logger.info("启用无头模式")

        chrome_options.add_argument(f"--window-size={self.window_size}")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        # 禁用自动化提示
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        try:
            service = Service(str(driver_path))
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            logger.info("ChromeDriver 启动成功")
            return self.driver

        except Exception as e:
            logger.error(f"启动 ChromeDriver 失败: {e}")
            raise

    def quit_driver(self) -> None:
        """关闭浏览器驱动"""
        if self.driver:
            try:
                self.driver.quit()
                logger.info("浏览器已关闭")
            except Exception as e:
                logger.error(f"关闭浏览器失败: {e}")
            finally:
                self.driver = None

    def get_driver(self) -> webdriver.Chrome:
        """
        获取驱动实例

        Returns:
            WebDriver 实例，如果不存在则创建新的
        """
        if self.driver is None:
            return self.create_driver()
        return self.driver
