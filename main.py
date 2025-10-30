#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化表单填写主程序
"""

import sys
import logging
import yaml
import time
from pathlib import Path
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# 添加 src 到路径
sys.path.insert(0, str(Path(__file__).parent / "src"))

from driver_manager import DriverManager
from data_reader import DataReader
from form_filler import FormFiller
from result_exporter import ResultExporter
from login_handler import LoginHandler


def setup_logging(config: dict) -> None:
    """配置日志"""
    log_config = config.get("logging", {})
    log_level = log_config.get("level", "INFO")
    log_file = log_config.get("file", "logs/app.log")
    log_format = log_config.get("format", "%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    # 确保日志目录存在
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)

    # 配置日志
    logging.basicConfig(
        level=getattr(logging, log_level),
        format=log_format,
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )


def load_config(config_path: str = "config/config.yaml") -> dict:
    """加载配置文件"""
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def handle_confirmation(driver, confirmation_config: dict, timeout: int = 30) -> None:
    """
    处理登录后的确认提示（如弹窗、对话框等）

    Args:
        driver: WebDriver 实例
        confirmation_config: 确认提示配置
        timeout: 超时时间（秒）
    """
    logger = logging.getLogger(__name__)

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

    try:
        confirmation_type = confirmation_config.get("type", "element")
        wait_time = confirmation_config.get("wait_time", 2)

        logger.info("等待确认提示出现...")
        time.sleep(wait_time)

        if confirmation_type == "alert":
            # 处理浏览器原生 alert 弹窗
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                logger.info(f"检测到 Alert 弹窗: {alert_text}")
                alert.accept()  # 点击确认
                logger.info("已点击 Alert 确认按钮")
            except Exception as e:
                logger.debug(f"未检测到 Alert 弹窗: {e}")

        elif confirmation_type == "element":
            # 处理页面元素类型的确认按钮
            button_config = confirmation_config.get("button", {})
            locator_type = button_config.get("locator", "id")
            locator_value = button_config.get("value")

            if not locator_value:
                logger.warning("确认按钮配置不完整，跳过")
                return

            by = LOCATOR_MAP.get(locator_type, By.ID)
            wait = WebDriverWait(driver, timeout)

            try:
                logger.info(f"查找确认按钮: {locator_type}={locator_value}")

                # 等待按钮元素存在
                confirm_button = wait.until(
                    EC.presence_of_element_located((by, locator_value))
                )
                logger.debug("确认按钮元素已找到")

                # 等待按钮可见
                wait.until(
                    EC.visibility_of_element_located((by, locator_value))
                )
                logger.debug("确认按钮已可见")

                # 直接使用 JavaScript 点击（避免元素遮挡问题）
                logger.info("使用 JavaScript 点击确认按钮...")
                driver.execute_script("arguments[0].click();", confirm_button)
                logger.info("已点击确认按钮")
                time.sleep(1)  # 等待确认完成

            except TimeoutException:
                logger.warning(f"超时：未找到确认按钮 (定位器: {locator_type}={locator_value})")
                logger.debug(f"当前页面URL: {driver.current_url}")

            except Exception as e:
                logger.warning(f"点击确认按钮失败: {e}")
                # 最后尝试：查找所有匹配的元素并逐个尝试点击
                try:
                    logger.info("尝试查找所有匹配元素...")
                    elements = driver.find_elements(by, locator_value)
                    logger.info(f"找到 {len(elements)} 个匹配元素")

                    for i, elem in enumerate(elements):
                        try:
                            if elem.is_displayed():
                                logger.info(f"尝试点击第 {i+1} 个元素...")
                                driver.execute_script("arguments[0].click();", elem)
                                logger.info(f"成功点击第 {i+1} 个元素")
                                time.sleep(1)
                                break
                        except Exception as elem_error:
                            logger.debug(f"点击第 {i+1} 个元素失败: {elem_error}")
                            continue
                except Exception as final_error:
                    logger.warning(f"所有尝试均失败: {final_error}")

        else:
            logger.warning(f"未知的确认类型: {confirmation_type}")

    except Exception as e:
        logger.error(f"处理确认提示时发生错误: {e}")
        # 不抛出异常，允许程序继续执行


def select_month(driver, month_config: dict, timeout: int = 30) -> None:
    """
    选择上报月份

    Args:
        driver: WebDriver 实例
        month_config: 月份选择配置
        timeout: 超时时间（秒）
    """
    logger = logging.getLogger(__name__)

    # 定位方式映射
    LOCATOR_MAP = {
        "id": By.ID,
        "name": By.NAME,
        "xpath": By.XPATH,
        "css_selector": By.CSS_SELECTOR,
        "class_name": By.CLASS_NAME,
    }

    try:
        # 获取月份配置
        month_value = month_config.get("month", "current")
        element_config = month_config.get("element", {})

        # 计算实际月份
        if month_value == "current":
            # 使用当前月份
            target_month = datetime.now().strftime("%Y-%m")
            logger.info(f"使用当前月份: {target_month}")
        else:
            # 使用配置的月份
            target_month = month_value
            logger.info(f"使用配置的月份: {target_month}")

        # 获取元素定位信息
        locator_type = element_config.get("locator", "id")
        locator_value = element_config.get("value")

        if not locator_value:
            logger.warning("月份输入框配置不完整，跳过月份选择")
            return

        by = LOCATOR_MAP.get(locator_type, By.ID)
        wait = WebDriverWait(driver, timeout)

        # 查找月份输入框
        logger.info(f"查找月份输入框: {locator_type}={locator_value}")
        month_input = wait.until(
            EC.presence_of_element_located((by, locator_value))
        )
        logger.debug("月份输入框已找到")

        # 使用 JavaScript 设置值（因为是 readonly 的 input）
        logger.info(f"设置月份为: {target_month}")
        driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input')); arguments[0].dispatchEvent(new Event('change'));",
            month_input,
            target_month
        )

        time.sleep(0.5)  # 等待值设置完成
        logger.info("月份值设置完成")

        # 点击确认按钮
        confirm_button_config = month_config.get("confirm_button", {})
        if confirm_button_config:
            button_locator_type = confirm_button_config.get("locator", "xpath")
            button_locator_value = confirm_button_config.get("value")

            if button_locator_value:
                button_by = LOCATOR_MAP.get(button_locator_type, By.XPATH)
                logger.info(f"查找月份确认按钮: {button_locator_type}={button_locator_value}")

                try:
                    confirm_button = wait.until(
                        EC.presence_of_element_located((button_by, button_locator_value))
                    )
                    logger.debug("确认按钮已找到")

                    # 使用 JavaScript 点击（避免元素遮挡问题）
                    driver.execute_script("arguments[0].click();", confirm_button)
                    logger.info("已点击月份确认按钮")
                    time.sleep(1)  # 等待 alert 弹窗出现

                    # 处理 alert 弹窗
                    try:
                        logger.info("等待 Alert 弹窗...")
                        alert = driver.switch_to.alert
                        alert_text = alert.text
                        logger.info(f"检测到 Alert 弹窗: {alert_text}")
                        alert.accept()  # 点击确认
                        logger.info("已点击 Alert 确认按钮")
                        time.sleep(0.5)  # 等待 alert 关闭
                    except Exception as alert_error:
                        logger.debug(f"未检测到 Alert 弹窗或处理失败: {alert_error}")

                except TimeoutException:
                    logger.warning(f"超时：未找到月份确认按钮 (定位器: {button_locator_type}={button_locator_value})")
                except Exception as btn_error:
                    logger.warning(f"点击月份确认按钮失败: {btn_error}")
            else:
                logger.debug("未配置确认按钮定位值，跳过")
        else:
            logger.debug("未配置月份确认按钮，跳过")

        logger.info("月份选择流程完成")

    except TimeoutException:
        logger.error(f"超时：找不到月份输入框 (定位器: {locator_type}={locator_value})")
    except Exception as e:
        logger.error(f"选择月份失败: {e}")
        # 不抛出异常，允许程序继续执行


def click_entry_button(driver, entry_button_config: dict, timeout: int = 30) -> None:
    """
    点击录入按钮

    Args:
        driver: WebDriver 实例
        entry_button_config: 录入按钮配置
        timeout: 超时时间（秒）
    """
    logger = logging.getLogger(__name__)

    # 定位方式映射
    LOCATOR_MAP = {
        "id": By.ID,
        "name": By.NAME,
        "xpath": By.XPATH,
        "css_selector": By.CSS_SELECTOR,
        "class_name": By.CLASS_NAME,
        "tag_name": By.TAG_NAME,
    }

    try:
        wait_time = entry_button_config.get("wait_time", 2)
        logger.info(f"等待录入按钮出现（{wait_time}秒）...")
        time.sleep(wait_time)

        # 获取按钮定位信息
        locator_type = entry_button_config.get("locator", "xpath")
        locator_value = entry_button_config.get("value")

        if not locator_value:
            logger.warning("录入按钮配置不完整，跳过")
            return

        by = LOCATOR_MAP.get(locator_type, By.XPATH)
        wait = WebDriverWait(driver, timeout)

        # 查找录入按钮
        logger.info(f"查找录入按钮: {locator_type}={locator_value}")
        entry_button = wait.until(
            EC.presence_of_element_located((by, locator_value))
        )
        logger.debug("录入按钮已找到")

        # 等待按钮可见
        wait.until(
            EC.visibility_of_element_located((by, locator_value))
        )
        logger.debug("录入按钮已可见")

        # 使用 JavaScript 点击（避免元素遮挡问题）
        logger.info("点击录入按钮...")
        driver.execute_script("arguments[0].click();", entry_button)
        logger.info("已点击录入按钮")
        time.sleep(1)  # 等待页面响应

    except TimeoutException:
        logger.error(f"超时：找不到录入按钮 (定位器: {locator_type}={locator_value})")
    except Exception as e:
        logger.error(f"点击录入按钮失败: {e}")
        # 不抛出异常，允许程序继续执行


def click_function_button(driver, function_button_config: dict, timeout: int = 30) -> None:
    """
    点击具体功能按钮（门诊/急诊）

    Args:
        driver: WebDriver 实例
        function_button_config: 功能按钮配置
        timeout: 超时时间（秒）
    """
    logger = logging.getLogger(__name__)

    # 定位方式映射
    LOCATOR_MAP = {
        "id": By.ID,
        "name": By.NAME,
        "xpath": By.XPATH,
        "css_selector": By.CSS_SELECTOR,
        "class_name": By.CLASS_NAME,
        "tag_name": By.TAG_NAME,
    }

    try:
        wait_time = function_button_config.get("wait_time", 2)
        logger.info(f"等待功能按钮出现（{wait_time}秒）...")
        time.sleep(wait_time)

        # 获取功能类型
        function_type = function_button_config.get("type", "outpatient")
        logger.info(f"选择的功能类型: {function_type}")

        # 获取对应类型的按钮配置
        button_config = function_button_config.get(function_type, {})
        if not button_config:
            logger.error(f"未找到功能类型 '{function_type}' 的配置")
            return

        # 获取按钮定位信息
        locator_type = button_config.get("locator", "xpath")
        locator_value = button_config.get("value")

        if not locator_value:
            logger.warning("功能按钮配置不完整，跳过")
            return

        by = LOCATOR_MAP.get(locator_type, By.XPATH)
        wait = WebDriverWait(driver, timeout)

        # 查找功能按钮
        logger.info(f"查找功能按钮: {locator_type}={locator_value}")
        function_button = wait.until(
            EC.presence_of_element_located((by, locator_value))
        )
        logger.debug("功能按钮已找到")

        # 等待按钮可见
        wait.until(
            EC.visibility_of_element_located((by, locator_value))
        )
        logger.debug("功能按钮已可见")

        # 使用 JavaScript 点击（避免元素遮挡问题）
        logger.info(f"点击功能按钮 ({function_type})...")
        driver.execute_script("arguments[0].click();", function_button)
        logger.info(f"已点击功能按钮 ({function_type})")
        time.sleep(1)  # 等待页面响应

    except TimeoutException:
        logger.error(f"超时：找不到功能按钮 (定位器: {locator_type}={locator_value})")
    except Exception as e:
        logger.error(f"点击功能按钮失败: {e}")
        # 不抛出异常，允许程序继续执行


def main():
    """主函数"""
    logger = logging.getLogger(__name__)

    try:
        # 加载配置
        logger.info("=" * 60)
        logger.info("自动化表单填写程序启动")
        logger.info("=" * 60)

        config = load_config()
        setup_logging(config)

        # 初始化组件
        logger.info("初始化组件...")

        # 1. 数据读取器
        data_config = config.get("data", {})
        reader = DataReader(
            file_path=data_config.get("input_file"),
            sheet_name=data_config.get("sheet_name", "Sheet1")
        )

        # 2. 结果导出器
        exporter = ResultExporter(
            output_file=data_config.get("output_file")
        )

        # 3. 浏览器驱动管理器
        browser_config = config.get("browser", {})
        driver_manager = DriverManager(
            headless=browser_config.get("headless", False),
            window_size=browser_config.get("window_size", "1920,1080")
        )

        # 4. 创建驱动
        driver = driver_manager.create_driver()

        # 5. 登录
        login_config = config.get("login", {})
        logger.info("=" * 60)
        logger.info("开始登录流程")
        logger.info("=" * 60)

        login_handler = LoginHandler(
            driver=driver,
            login_config=login_config,
            timeout=browser_config.get("timeout", 30)
        )
        login_success = login_handler.login()

        if not login_success:
            raise Exception("登录失败，程序终止")

        logger.info("=" * 60)
        logger.info("登录流程完成")
        logger.info("=" * 60)

        # 6. 处理登录后的确认提示
        confirmation_config = login_config.get("confirmation", {})
        logger.info("=" * 60)
        logger.info("处理登录后的确认提示")
        logger.info("=" * 60)

        handle_confirmation(
            driver=driver,
            confirmation_config=confirmation_config,
            timeout=browser_config.get("timeout", 30)
        )

        logger.info("=" * 60)
        logger.info("确认提示处理完成")
        logger.info("=" * 60)

        # 7. 选择月份
        month_config = config.get("month_selection", {})
        logger.info("=" * 60)
        logger.info("选择上报月份")
        logger.info("=" * 60)

        select_month(
            driver=driver,
            month_config=month_config,
            timeout=browser_config.get("timeout", 30)
        )

        logger.info("=" * 60)
        logger.info("月份选择完成")
        logger.info("=" * 60)

        # 8. 点击录入按钮
        entry_button_config = config.get("entry_button", {})
        logger.info("=" * 60)
        logger.info("点击录入按钮")
        logger.info("=" * 60)

        click_entry_button(
            driver=driver,
            entry_button_config=entry_button_config,
            timeout=browser_config.get("timeout", 30)
        )

        logger.info("=" * 60)
        logger.info("录入按钮点击完成")
        logger.info("=" * 60)

        # 9. 点击功能按钮（门诊/急诊）
        function_button_config = config.get("function_button", {})
        logger.info("=" * 60)
        logger.info("点击功能按钮")
        logger.info("=" * 60)

        click_function_button(
            driver=driver,
            function_button_config=function_button_config,
            timeout=browser_config.get("timeout", 30)
        )

        logger.info("=" * 60)
        logger.info("功能按钮点击完成")
        logger.info("=" * 60)

        # 10. 表单填写器
        form_filler = FormFiller(
            driver=driver,
            form_elements=config.get("form_elements", {}),
            timeout=browser_config.get("timeout", 30)
        )

        # 读取数据
        logger.info("读取输入数据...")
        data_list = reader.read_data()
        total_count = len(data_list)
        logger.info(f"共读取 {total_count} 条数据")

        # 处理每条数据
        results = []
        success_count = 0
        fail_count = 0

        for index, row_data in enumerate(data_list, start=1):
            logger.info(f"\n处理第 {index}/{total_count} 条数据...")

            try:
                # 填写表单
                fill_success = form_filler.fill_form(row_data)

                if not fill_success:
                    raise Exception("表单填写失败")

                # 提交表单
                submit_success = form_filler.submit_form()

                if not submit_success:
                    raise Exception("表单提交失败")

                # 记录成功结果
                result = exporter.create_result_entry(
                    row_data=row_data,
                    status="成功",
                    message="表单提交成功"
                )
                results.append(result)
                success_count += 1
                logger.info(f"第 {index} 条数据处理成功")

                # 刷新页面，准备处理下一条数据
                if index < total_count:
                    driver.refresh()
                    logger.info("页面已刷新，准备处理下一条数据")

            except Exception as e:
                # 记录失败结果
                result = exporter.create_result_entry(
                    row_data=row_data,
                    status="失败",
                    message=str(e)
                )
                results.append(result)
                fail_count += 1
                logger.error(f"第 {index} 条数据处理失败: {e}")

                # 失败后刷新页面
                try:
                    driver.refresh()
                except:
                    pass

        # 导出结果
        logger.info("\n" + "=" * 60)
        logger.info("处理完成，导出结果...")
        exporter.export_results(results)

        # 统计信息
        logger.info("=" * 60)
        logger.info(f"处理总数: {total_count}")
        logger.info(f"成功: {success_count}")
        logger.info(f"失败: {fail_count}")
        logger.info(f"成功率: {success_count / total_count * 100:.2f}%")
        logger.info("=" * 60)

    except KeyboardInterrupt:
        logger.warning("\n用户中断程序")

    except Exception as e:
        logger.error(f"程序执行出错: {e}", exc_info=True)

    finally:
        # 关闭浏览器
        try:
            driver_manager.quit_driver()
        except:
            pass

        logger.info("程序结束")


if __name__ == "__main__":
    main()
