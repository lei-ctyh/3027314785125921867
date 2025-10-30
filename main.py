#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化表单填写主程序
"""

import sys
import logging
import yaml
from pathlib import Path

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

        # 5. 登录处理（如果启用）
        login_config = config.get("login", {})
        if login_config.get("enabled", False):
            logger.info("=" * 60)
            logger.info("开始登录流程")
            logger.info("=" * 60)

            login_handler = LoginHandler(
                driver=driver,
                login_config=login_config,
                timeout=browser_config.get("timeout", 30)
            )

            # 执行登录
            login_success = login_handler.login()

            if not login_success:
                raise Exception("登录失败，程序终止")

            # 登录成功后导航到表单页面
            nav_success = login_handler.navigate_to_form_page()

            if not nav_success:
                raise Exception("导航到表单页面失败，程序终止")

            logger.info("=" * 60)
            logger.info("登录流程完成")
            logger.info("=" * 60)
        else:
            # 如果不需要登录，直接打开表单页面
            website_url = config.get("website", {}).get("url")
            logger.info(f"打开目标网站: {website_url}")
            driver.get(website_url)

        # 6. 表单填写器
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
