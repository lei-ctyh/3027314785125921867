"""
表单填写核心模块
"""

import logging
import time
import re
import requests
from typing import Dict, Any, List, Optional, Tuple
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

    def __init__(self, driver, form_elements: Dict, timeout: int = 30, antibiotic_config: Dict = None):
        """
        初始化表单填写器

        Args:
            driver: WebDriver 实例
            form_elements: 表单元素配置
            timeout: 超时时间（秒）
            antibiotic_config: 抗菌药处理配置
        """
        self.driver = driver
        self.form_elements = form_elements
        self.timeout = timeout
        self.wait = WebDriverWait(driver, timeout)
        self.session = requests.Session()  # 用于API请求
        self.antibiotic_config = antibiotic_config or {}

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
                field_name = field_name.split("\n")[0].strip()

                # 跳过空值
                if field_value is None or str(field_value).strip() == "":
                    logger.debug(f"跳过空字段: {field_name}")
                    continue

                # 获取元素配置
                element_config = self.form_elements.get(field_name)
                if not element_config and field_name != '诊断':
                    logger.warning(f"配置中未找到字段: {field_name}")
                    continue

                # 填写字段
                if field_name == "科室":
                    field_value = field_value.replace(" ", "").replace("门诊", "")
                    self._fill_field(field_name, field_value, element_config)
                elif field_name == "年龄":
                    field_value1 = field_value[-1]
                    field_value = field_value[0:-1]
                    element_config1 = self.form_elements.get("年龄单位")
                    self._fill_field("年龄单位", field_value1, element_config1)
                    self._fill_field(field_name, field_value, element_config)
                elif field_name == "药品品种数":
                    field_value = int(field_value)
                    self._fill_field(field_name, field_value, element_config)
                elif field_name == "注射剂":
                    self._fill_field(field_name, field_value, element_config)
                    if field_value == '有':
                        element_config1 = self.form_elements.get("注射剂数量")
                        self._fill_field("注射剂数量", "1", element_config1)
                elif field_name == "诊断":
                    field_values = field_value.replace("，",",").split(",")
                    for i, value in enumerate(field_values[:5]):
                        if not value.strip():
                            continue

                        # 查询诊断编码
                        diag_info = self._search_diagnosis(value.strip())
                        if diag_info:
                            # 填入诊断名称和编码
                            index = i + 1
                            name_config = self.form_elements.get(f"诊断{index}_名称")
                            code_config = self.form_elements.get(f"诊断{index}_编码")

                            if name_config:
                                self._fill_field(f"诊断{index}_名称", diag_info['name'], name_config)
                            if code_config:
                                self._fill_field(f"诊断{index}_编码", diag_info['code'], code_config)

                            logger.info(f"填入诊断{index}: {diag_info['name']} ({diag_info['code']})")
                        else:
                            logger.warning(f"未找到诊断信息: {value}")
                else:
                    self._fill_field(field_name, field_value, element_config)


                time.sleep(0.5)  # 短暂延迟，模拟人工操作

            logger.info("表单填写完成")

            # 填写成功，点击提交按钮
            if self.submit_form("submit_button"):
                return True
            else:
                logger.error("提交失败，点击重置按钮")
                self._click_button("reset_button")
                return False

        except Exception as e:
            logger.error(f"填写表单失败: {e}")
            # 填写失败，点击重置按钮
            self._click_button("reset_button")
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
                # 对于诊断名称和编码，使用JavaScript直接设置值，避免触发onfocus事件
                if "诊断" in field_name and ("名称" in field_name or "编码" in field_name):
                    # 先移除onfocus和onblur事件
                    self.driver.execute_script(
                        "arguments[0].removeAttribute('onfocus');"
                        "arguments[0].removeAttribute('onblur');",
                        element
                    )
                    # 使用JavaScript设置值
                    self.driver.execute_script(
                        "arguments[0].value = arguments[1];",
                        element, str(field_value)
                    )
                    logger.debug(f"使用JS填写诊断字段 {field_name}: {field_value}")
                else:
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

            elif element_type == "radio":
                # 处理单选框：根据配置的options找到对应的value
                options = config.get("options", {})
                radio_value = options.get(str(field_value), str(field_value))

                # 构造带value的选择器
                if locator_type == "name":
                    # 对于name类型，使用CSS选择器查找特定value的radio
                    radio_by = By.CSS_SELECTOR
                    radio_locator = f"input[name='{locator_value}'][value='{radio_value}']"
                elif locator_type == "id":
                    # 对于id类型，直接使用id（需要配置中指定完整的id）
                    radio_by = By.ID
                    radio_locator = f"{locator_value}"
                else:
                    radio_by = by
                    radio_locator = locator_value

                radio_element = self.driver.find_element(radio_by, radio_locator)
                if not radio_element.is_selected():
                    radio_element.click()
                logger.debug(f"选择单选框 {field_name}: {field_value} (value={radio_value})")

            elif element_type == "button":
                element.click()
                logger.debug(f"点击按钮 {field_name}")

            elif element_type == "hidden":
                # 隐藏字段不需要填写，跳过
                logger.debug(f"跳过隐藏字段 {field_name}")

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
            logger.info("已点击提交按钮")
            time.sleep(1)  # 等待 alert 弹出

            # 处理 alert 弹窗
            try:
                logger.info("等待 Alert 弹窗...")
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                logger.info(f"检测到 Alert 弹窗: {alert_text}")
                alert.accept()  # 点击确认
                logger.info("已点击 Alert 确认按钮")
                time.sleep(1)  # 等待 alert 关闭
            except Exception as alert_error:
                logger.debug(f"未检测到 Alert 弹窗或处理失败: {alert_error}")

            logger.info("表单提交成功")
            return True

        except Exception as e:
            logger.error(f"提交表单失败: {e}")
            return False

    def _click_button(self, button_name: str) -> bool:
        """
        点击按钮（通用方法）

        Args:
            button_name: 按钮在配置中的名称

        Returns:
            True 如果点击成功
        """
        try:
            button_config = self.form_elements.get(button_name)
            if not button_config:
                logger.warning(f"配置中未找到按钮: {button_name}")
                return False

            locator_type = button_config.get("locator", "id")
            locator_value = button_config.get("value")
            by = self.LOCATOR_MAP.get(locator_type, By.ID)

            # 等待按钮可点击
            button = self.wait.until(
                EC.element_to_be_clickable((by, locator_value))
            )

            button.click()
            logger.info(f"点击按钮 {button_name} 成功")
            time.sleep(1)  # 短暂等待
            return True

        except Exception as e:
            logger.warning(f"点击按钮 {button_name} 失败: {e}")
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

    def _search_diagnosis(self, keyword: str) -> Optional[Dict[str, str]]:
        """
        查询诊断编码和名称

        Args:
            keyword: 诊断关键词

        Returns:
            包含诊断名称和编码的字典，格式: {'name': '诊断名称', 'code': '编码'}
            如果未找到则返回 None
        """
        try:
            # 从浏览器获取Cookie
            cookies = self.driver.get_cookies()
            cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

            # 构造请求
            url = "http://y.chinadtc.org.cn/entering/dict/search_dict"
            headers = {
                'Accept': '*/*',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
                'Cookie': 'PHPSESSID=' + cookie_dict['PHPSESSID']
            }

            # 构造表单数据（multipart/form-data格式）
            data = {
                'dict_table': 'dict_diag',
                'search_field': 'diag_pym',
                'order_field': 'diag_id',
                'szimu': keyword
            }

            # 将data转换成multipart形式
            # 每个字段作为元组，第二个元素为None表示不是文件
            files = []
            for key, value in data.items():
                files.append((key, (None, value)))

            # 发送请求，不设置Content-Type，让requests自动设置（包含boundary）
            response = self.session.post(url, files=files, headers=headers, timeout=10)
            response.raise_for_status()

            # 解析响应
            results = response.json()

            if not results or len(results) == 0:
                logger.warning(f"未找到诊断: {keyword}")
                return None

            # 选择最相似的结果（这里简单取第一个）
            # 可以实现更复杂的相似度匹配算法
            best_match = results[0]

            diag_info = {
                'name': best_match['diag_name'],
                'code': best_match['diag_code']
            }

            logger.info(f"找到诊断: {keyword} -> {diag_info['name']} ({diag_info['code']})")
            return diag_info

        except requests.RequestException as e:
            logger.error(f"查询诊断API失败: {e}")
            return None
        except (KeyError, IndexError, ValueError) as e:
            logger.error(f"解析诊断结果失败: {e}")
            return None
        except Exception as e:
            logger.error(f"查询诊断失败: {e}")
            return None

    def handle_antibiotic_info(self, row_data: Dict[str, Any]) -> bool:
        """
        处理新增记录的抗菌药信息

        Args:
            row_data: 当前行的数据字典

        Returns:
            True 如果处理成功，False 否则
        """
        try:
            # 检查是否启用抗菌药处理
            if not self.antibiotic_config.get("enabled", False):
                logger.info("抗菌药处理未启用，跳过")
                return True

            # 获取抗菌药有/无的值（可能带换行符）
            antibiotic_value = None
            for key in row_data.keys():
                if '抗菌药' in key and ('有' in key or '无' in key):
                    antibiotic_value = str(row_data[key]).strip()
                    break

            if not antibiotic_value or antibiotic_value == '':
                logger.info("未找到抗菌药字段或值为空，跳过抗菌药处理")
                return True

            logger.info(f"开始处理抗菌药信息，值: {antibiotic_value}")

            # 从配置中获取参数
            table_id = self.antibiotic_config.get("result_table_id", "outpatientTable")
            wait_time = self.antibiotic_config.get("wait_after_submit", 2)

            # 等待表格刷新（等待新记录添加到表格中）
            time.sleep(wait_time)

            # 查找表格
            table = self.wait.until(
                EC.presence_of_element_located((By.ID, table_id))
            )
            logger.debug(f"找到结果表格: {table_id}")

            # 查找表格的第一个数据行（跳过表头）
            # 表头的class是"tabletitle"，所以我们找第一个没有这个class的tr
            rows = table.find_elements(By.TAG_NAME, "tr")
            first_data_row = None
            for row in rows:
                row_class = row.get_attribute("class") or ""
                if "tabletitle" not in row_class and row.get_attribute("id"):
                    first_data_row = row
                    break

            if not first_data_row:
                logger.error("未找到表格的第一行数据")
                return False

            row_id = first_data_row.get_attribute("id")
            logger.info(f"找到新增记录行: {row_id}")

            # 在这一行中查找抗菌药的单选按钮
            # 先找到所有的radio按钮，通过name属性识别（name="drugsMoney{序号}"）
            radio_buttons = first_data_row.find_elements(By.CSS_SELECTOR, "input[type='radio'][name^='drugsMoney']")

            if len(radio_buttons) < 2:
                logger.error(f"未找到足够的抗菌药单选按钮，找到 {len(radio_buttons)} 个")
                return False

            # 找出"有"和"无"的按钮
            radio_no = None  # value="0"
            radio_yes = None  # value="1"

            for radio in radio_buttons:
                value = radio.get_attribute("value")
                if value == "0":
                    radio_no = radio
                elif value == "1":
                    radio_yes = radio

            # 根据数据值选择按钮
            if antibiotic_value == "有":
                if radio_yes:
                    logger.info("选择抗菌药: 有")
                    # 使用JavaScript点击，避免遮挡问题
                    self.driver.execute_script("arguments[0].click();", radio_yes)
                    time.sleep(1)  # 等待按钮启用

                    # 查找"录入详细信息"按钮
                    detail_button = first_data_row.find_element(By.CSS_SELECTOR, "input.itemBtnDrugs.btnDrugs")

                    # 检查按钮是否已启用
                    is_disabled = detail_button.get_attribute("disabled")
                    if is_disabled:
                        logger.warning("录入详细信息按钮仍处于禁用状态")
                        # 再等一会儿
                        time.sleep(1)

                    # 点击"录入详细信息"按钮
                    logger.info("点击录入详细信息按钮")
                    self.driver.execute_script("arguments[0].click();", detail_button)

                    # 调用详情填写方法
                    logger.info("准备填写抗菌药详细信息...")
                    detail_success = self.fill_antibiotic_detail(row_data)

                    if not detail_success:
                        logger.warning("抗菌药详细信息填写失败")
                        return False

                    logger.info("抗菌药信息处理完成：已选择'有'并完成详细信息录入")
                    return True
                else:
                    logger.error("未找到'有'的单选按钮")
                    return False

            elif antibiotic_value == "无":
                if radio_no:
                    logger.info("选择抗菌药: 无")
                    # 默认已经选中"无"，但为了确保，还是点击一下
                    self.driver.execute_script("arguments[0].click();", radio_no)
                    time.sleep(0.5)
                    logger.info("抗菌药信息处理完成：已选择'无'")
                    return True
                else:
                    logger.warning("未找到'无'的单选按钮，使用默认值")
                    return True

            else:
                logger.warning(f"未知的抗菌药值: {antibiotic_value}，跳过处理")
                return True

        except TimeoutException:
            logger.error(f"超时：未找到结果表格 (ID: {self.antibiotic_config.get('result_table_id', 'outpatientTable')})")
            return False
        except NoSuchElementException as e:
            logger.error(f"未找到元素: {e}")
            return False
        except Exception as e:
            logger.error(f"处理抗菌药信息失败: {e}")
            return False

    def _search_drug(self, keyword: str) -> Optional[Dict[str, str]]:
        """
        查询药品通用名和编码

        Args:
            keyword: 药品关键词

        Returns:
            包含药品名称、编码和规格的字典，格式: {'name': '药品名称', 'code': '编码', 'spec': '规格'}
            如果未找到则返回 None
        """
        try:
            # 从浏览器获取Cookie
            cookies = self.driver.get_cookies()
            cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}

            # 构造请求
            url = "http://y.chinadtc.org.cn/entering/dict/search_dict"
            headers = {
                'Accept': '*/*',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
                'Cookie': 'PHPSESSID=' + cookie_dict['PHPSESSID']
            }

            # 构造表单数据（multipart/form-data格式）
            data = {
                'dict_table': 'dict_drug',
                'search_field': 'drug_pym',
                'order_field': 'drug_id',
                'szimu': keyword
            }

            # 将data转换成multipart形式
            files = []
            for key, value in data.items():
                files.append((key, (None, value)))

            # 发送请求，不设置Content-Type，让requests自动设置（包含boundary）
            response = self.session.post(url, files=files, headers=headers, timeout=10)
            response.raise_for_status()

            # 解析响应
            results = response.json()

            if not results or len(results) == 0:
                logger.warning(f"未找到药品: {keyword}")
                return None

            # 选择最相似的结果（这里简单取第一个）
            best_match = results[0]

            # 构造规格字符串
            spec = best_match.get('drug_spec_c', '')
            spec_unit2 = best_match.get('drug_spec_unit2', '')
            drug_form = best_match.get('drug_form_c', '')

            # 组合规格，例如: "0.25g 胶囊"
            full_spec = f"{spec}{spec_unit2}" if spec else ""
            if drug_form:
                full_spec = f"{full_spec} {drug_form}" if full_spec else drug_form

            drug_info = {
                'name': best_match['drug_name'],
                'code': best_match['drug_code'],
                'spec': full_spec,
                'id': str(best_match['drug_id'])
            }

            logger.info(f"找到药品: {keyword} -> {drug_info['name']} ({drug_info['code']}) 规格: {drug_info['spec']}")
            return drug_info

        except requests.RequestException as e:
            logger.error(f"查询药品API失败: {e}")
            return None
        except (KeyError, IndexError, ValueError) as e:
            logger.error(f"解析药品结果失败: {e}")
            return None
        except Exception as e:
            logger.error(f"查询药品失败: {e}")
            return None

    def _parse_dosage(self, dosage_str: str) -> Dict[str, Any]:
        """
        解析用法用量字符串

        Args:
            dosage_str: 用法用量字符串，如 "100mg bid", "0.25g tid", "125mg q8h"

        Returns:
            包含单次剂量、单次剂量单位、用法频率的字典
            格式: {'dose_value': '100', 'dose_unit': 'mg', 'frequency': 'bid'}
        """
        try:
            logger.debug(f"解析用法用量: {dosage_str}")

            # 初始化默认值
            result = {
                'dose_value': '',
                'dose_unit': '',
                'frequency': '',
                'total_amount': '',  # 总量（从数量列获取）
                'total_unit': ''     # 总量单位
            }

            if not dosage_str or str(dosage_str).strip() == '':
                logger.warning("用法用量字符串为空")
                return result

            dosage_str = str(dosage_str).strip()

            # 正则表达式匹配剂量和单位，例如: "100mg", "0.25g", "1片"
            dose_pattern = r'(\d+\.?\d*)\s*(mg|g|克|毫克|片|粒|支|包|袋|瓶|ml|毫升|滴|万单位)'
            dose_match = re.search(dose_pattern, dosage_str, re.IGNORECASE)

            if dose_match:
                result['dose_value'] = dose_match.group(1)
                result['dose_unit'] = dose_match.group(2).lower()
                logger.debug(f"提取到剂量: {result['dose_value']} {result['dose_unit']}")

            # 正则表达式匹配用法频率
            # 支持: qd, bid, tid, qid, q2h, q4h, q6h, q8h, q12h, qn (每晚)
            frequency_pattern = r'\b(qd|bid|tid|qid|q2h|q4h|q6h|q8h|q12h|qn|st|即刻|1日|2日|3日|4日|每晚)\b'
            freq_match = re.search(frequency_pattern, dosage_str, re.IGNORECASE)

            if freq_match:
                result['frequency'] = freq_match.group(1).lower()
                logger.debug(f"提取到频率: {result['frequency']}")

            return result

        except Exception as e:
            logger.error(f"解析用法用量失败: {e}")
            return {
                'dose_value': '',
                'dose_unit': '',
                'frequency': '',
                'total_amount': '',
                'total_unit': ''
            }

    def _normalize_unit(self, unit_str: str, unit_type: str = 'dose') -> str:
        """
        标准化单位值到HTML select的option value

        Args:
            unit_str: 单位字符串 (如 "mg", "g", "片", "口服")
            unit_type: 单位类型 ('dose' 剂量单位, 'route' 给药途径)

        Returns:
            HTML select对应的value值 (字符串)
        """
        unit_str = str(unit_str).strip().lower()

        if unit_type == 'dose':
            # 剂量单位映射表
            unit_map = {
                '克': '5', 'g': '5',
                '毫克': '6', 'mg': '6',
                '万单位': '7',
                '滴': '8',
                'ml': '9', '毫升': '9',
                '片': '10',
                '支': '11',
                '粒': '12',
                '瓶': '13',
                '包': '14',
                '袋': '15'
            }
            return unit_map.get(unit_str, '10')  # 默认"片"

        elif unit_type == 'route':
            # 给药途径映射表
            route_map = {
                '静脉滴注': '26', '静滴': '26',
                '静脉泵入': '27',
                '静脉推注': '28', '静推': '28',
                '肌肉注射': '29', '肌注': '29',
                '静脉注射': '30', '静注': '30',
                '皮下注射': '31', '皮下': '31',
                '球后注射': '87',
                '结膜下注射': '88',
                '眼内注射': '89',
                '直肠给药': '32', '直肠': '32',
                '雾化吸入': '33', '雾化': '33',
                '肠道准备': '34',
                '口服': '35',
                '外用': '36',
                '滴鼻': '37',
                '滴耳': '38',
                '滴眼': '39',
                '鞘内注射': '90',
                '腹膜透析': '91',
                '皮试': '92'
            }
            return route_map.get(unit_str, '35')  # 默认"口服"

        elif unit_type == 'frequency':
            # 用法频率映射表
            freq_map = {
                '即刻': '16', 'st': '16',
                '1/日': '17', 'qd': '17', '1日': '17',
                '2/日': '18', 'bid': '18', '2日': '18',
                '3/日': '19', 'tid': '19', '3日': '19',
                '4/日': '20', 'qid': '20', '4日': '20',
                'q2h': '93',
                'q6h': '21',
                'q8h': '22',
                'q12h': '23',
                '每晚': '24', 'qn': '24',
                '其他': '25'
            }
            return freq_map.get(unit_str, '17')  # 默认"1/日"

        return ''

    def fill_antibiotic_detail(self, row_data: Dict[str, Any]) -> bool:
        """
        填写抗菌药详细信息表单

        Args:
            row_data: 当前行的数据字典（包含抗菌药相关字段）

        Returns:
            True 如果填写成功，False 否则
        """
        try:
            logger.info("开始填写抗菌药详细信息...")

            # 等待抗菌药详情页面加载
            wait_time = self.antibiotic_config.get('antibiotic_detail', {}).get('wait_after_click', 2)
            logger.info(f"等待页面加载（{wait_time}秒）...")
            time.sleep(wait_time)

            # 检查页面是否加载完成（检测特征元素）
            try:
                self.wait.until(
                    EC.presence_of_element_located((By.ID, "drug_idName"))
                )
                logger.debug("抗菌药详情页面已加载")
            except TimeoutException:
                logger.error("抗菌药详情页面加载超时")
                return False

            # 提取Excel中的相关字段数据
            drug_name_raw = None
            drug_spec = None
            drug_amount = None
            drug_dosage = None
            drug_route = None
            drug_quantity = None

            # 遍历row_data查找相关字段（处理列名可能的换行符和空格）
            for key, value in row_data.items():
                key_clean = key.replace('\n', '').strip()
                if '药品名称' in key_clean or '药品' in key_clean:
                    drug_name_raw = str(value).strip() if value else None
                elif '规格' in key_clean:
                    drug_spec = str(value).strip() if value else None
                elif '金额' in key_clean and '元' in key_clean and '处方' not in key_clean:
                    drug_amount = str(value).strip() if value else None
                elif '用法用量' in key_clean or '用法' in key_clean:
                    drug_dosage = str(value).strip() if value else None
                elif '途径' in key_clean:
                    drug_route = str(value).strip() if value else None
                elif '数量' in key_clean:
                    drug_quantity = str(value).strip() if value else None

            logger.info(f"提取到的数据 - 药品:{drug_name_raw}, 规格:{drug_spec}, 金额:{drug_amount}, 用法用量:{drug_dosage}, 途径:{drug_route}, 数量:{drug_quantity}")

            # 1. 查询药品通用名
            drug_info = None
            if drug_name_raw:
                drug_info = self._search_drug(drug_name_raw)
                if drug_info:
                    # 填写药品通用名和规格
                    logger.info(f"填写药品通用名: {drug_info['name']}")
                    # 使用JavaScript直接设置值（readonly字段）
                    medicine_name_input = self.driver.find_element(By.ID, "medicineName")
                    self.driver.execute_script("arguments[0].value = arguments[1];", medicine_name_input, drug_info['name'])

                    # 填写药品ID（隐藏字段）
                    drug_id_input = self.driver.find_element(By.ID, "drug_idName")
                    self.driver.execute_script("arguments[0].value = arguments[1];", drug_id_input, drug_info['id'])

                    # 填写规格
                    spec_name_input = self.driver.find_element(By.ID, "specName")
                    spec_value = drug_info.get('spec', drug_spec or '')
                    self.driver.execute_script("arguments[0].value = arguments[1];", spec_name_input, spec_value)
                    logger.info(f"填写规格: {spec_value}")
                else:
                    logger.warning(f"未查询到药品通用名，使用原始名称: {drug_name_raw}")
                    medicine_name_input = self.driver.find_element(By.ID, "medicineName")
                    self.driver.execute_script("arguments[0].value = arguments[1];", medicine_name_input, drug_name_raw)
                    if drug_spec:
                        spec_name_input = self.driver.find_element(By.ID, "specName")
                        self.driver.execute_script("arguments[0].value = arguments[1];", spec_name_input, drug_spec)

            # 2. 填写金额
            if drug_amount:
                logger.info(f"填写金额: {drug_amount}")
                amount_input = self.driver.find_element(By.ID, "amountOutpatient")
                amount_input.clear()
                amount_input.send_keys(str(drug_amount))

            # 3. 解析用法用量
            dosage_info = {}
            if drug_dosage:
                dosage_info = self._parse_dosage(drug_dosage)
                logger.info(f"解析用法用量结果: {dosage_info}")

            # 4. 解析数量（总用量）
            # 数量格式如 "3.00个", "4.00盒"
            total_amount = ''
            total_unit = ''
            if drug_quantity:
                # 提取数字和单位
                qty_match = re.match(r'(\d+\.?\d*)\s*(个|盒|瓶|支|片|粒|包|袋|克|g|mg|毫克)?', str(drug_quantity))
                if qty_match:
                    total_amount = qty_match.group(1)
                    if qty_match.group(2):
                        total_unit = qty_match.group(2)

            # 如果没有从数量提取到，使用用法用量中的剂量
            if not total_amount and dosage_info.get('dose_value'):
                total_amount = dosage_info['dose_value']
                total_unit = dosage_info.get('dose_unit', '')

            # 5. 填写总用量
            if total_amount:
                logger.info(f"填写总用量: {total_amount}")
                total_medicine_input = self.driver.find_element(By.ID, "totalMedicine")
                total_medicine_input.clear()
                total_medicine_input.send_keys(str(total_amount))

            # 6. 选择总用量单位
            if total_unit:
                unit_value = self._normalize_unit(total_unit, 'dose')
                logger.info(f"选择总用量单位: {total_unit} -> value={unit_value}")
                total_unit_select = Select(self.driver.find_element(By.ID, "totalMedicineUnit"))
                total_unit_select.select_by_value(unit_value)

            # 7. 填写单次计量
            if dosage_info.get('dose_value'):
                logger.info(f"填写单次计量: {dosage_info['dose_value']}")
                once_meter_input = self.driver.find_element(By.ID, "onceMeter")
                once_meter_input.clear()
                once_meter_input.send_keys(str(dosage_info['dose_value']))

            # 8. 选择单次计量单位
            if dosage_info.get('dose_unit'):
                unit_value = self._normalize_unit(dosage_info['dose_unit'], 'dose')
                logger.info(f"选择单次计量单位: {dosage_info['dose_unit']} -> value={unit_value}")
                once_unit_select = Select(self.driver.find_element(By.ID, "onceMeterUnit"))
                once_unit_select.select_by_value(unit_value)

            # 9. 选择用法（频率）
            if dosage_info.get('frequency'):
                freq_value = self._normalize_unit(dosage_info['frequency'], 'frequency')
                logger.info(f"选择用法频率: {dosage_info['frequency']} -> value={freq_value}")
                freq_select = Select(self.driver.find_element(By.ID, "medicineFrequency"))
                freq_select.select_by_value(freq_value)

            # 10. 选择途径
            if drug_route:
                route_value = self._normalize_unit(drug_route, 'route')
                logger.info(f"选择途径: {drug_route} -> value={route_value}")
                route_select = Select(self.driver.find_element(By.ID, "medicineWay"))
                route_select.select_by_value(route_value)

            time.sleep(0.5)  # 短暂等待表单填写完成

            # 11. 点击保存按钮
            logger.info("点击保存按钮...")
            save_button = self.driver.find_element(By.XPATH, "//input[@value='保存抗菌药详细信息录入']")
            self.driver.execute_script("arguments[0].click();", save_button)
            time.sleep(1)  # 等待保存处理

            # 12. 处理alert弹窗
            try:
                logger.info("等待 Alert 弹窗...")
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                logger.info(f"检测到 Alert 弹窗: {alert_text}")
                alert.accept()  # 点击确认
                logger.info("已点击 Alert 确认按钮")
                time.sleep(0.5)  # 等待 alert 关闭
            except Exception as alert_error:
                logger.debug(f"未检测到 Alert 弹窗或处理失败: {alert_error}")

            # 13. 点击返回按钮
            logger.info("查找返回按钮...")
            try:
                # 尝试多种定位方式
                return_button = None
                try:
                    return_button = self.driver.find_element(By.XPATH, "//input[@value='返回门诊处方用药情况调查表']")
                except:
                    try:
                        return_button = self.driver.find_element(By.XPATH, "//input[@onclick=\"fanhui('1')\"]")
                    except:
                        logger.warning("未找到返回按钮，可能已自动返回")

                if return_button:
                    logger.info("点击返回按钮...")
                    self.driver.execute_script("arguments[0].click();", return_button)
                    time.sleep(1)  # 等待返回主列表
                    logger.info("已返回主列表")

            except Exception as return_error:
                logger.warning(f"点击返回按钮失败: {return_error}")

            logger.info("抗菌药详细信息填写完成")
            return True

        except NoSuchElementException as e:
            logger.error(f"未找到元素: {e}")
            return False
        except Exception as e:
            logger.error(f"填写抗菌药详细信息失败: {e}")
            return False
