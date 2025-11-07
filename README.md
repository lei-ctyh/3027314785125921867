# 自动化表单填写系统

一个基于 Selenium 的自动化表单填写工具，支持从 Excel/CSV 文件批量读取数据并自动填写网页表单。

## 功能特点

- **登录支持**：自动登录网站（支持固定账号密码）
- 从 Excel/CSV 文件批量读取表单数据
- 自动填写网页表单（支持文本框、下拉框、文本域等）
- 批量提交表单
- 自动导出处理结果到 Excel
- 详细的日志记录
- 灵活的配置文件，支持自定义表单元素定位
- 支持无头模式（后台运行）

## 项目结构

```
.
├── config/                 # 配置文件目录
│   └── config.yaml        # 主配置文件
├── data/                  # 数据文件目录
│   └── input_data.xlsx    # 输入数据（Excel格式）
├── output/                # 输出结果目录
│   └── results_*.xlsx     # 处理结果（自动生成）
├── logs/                  # 日志文件目录
│   └── app.log           # 运行日志
├── src/                   # 源代码目录
│   ├── __init__.py
│   ├── driver_manager.py  # 浏览器驱动管理
│   ├── data_reader.py     # 数据读取模块
│   ├── form_filler.py     # 表单填写核心模块
│   ├── login_handler.py   # 登录处理模块
│   └── result_exporter.py # 结果导出模块
├── chromedriver-win64/    # ChromeDriver 目录
│   └── chromedriver.exe
├── main.py                # 主程序入口
├── requirements.txt       # Python 依赖
└── README.md             # 项目说明文档
```

## 安装步骤

### 1. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 2. 下载 ChromeDriver

1. 查看你的 Chrome 浏览器版本（在浏览器地址栏输入 `chrome://version/`）
2. 下载对应版本的 ChromeDriver：
   - 下载地址：https://googlechromelabs.github.io/chrome-for-testing/
   - 选择 `chromedriver` -> `win64` -> 下载 zip 文件
3. 解压后将 `chromedriver.exe` 放到项目根目录的 `chromedriver-win64/` 文件夹中

## 配置说明

### 1. 修改配置文件 `config/config.yaml`

#### 1.1 登录配置（如果需要登录）

如果目标网站需要登录才能访问表单，请配置以下内容：

```yaml
login:
  enabled: true  # 是否需要登录（true=需要登录，false=不需要登录）
  login_url: "http://example.com/login"  # 登录页面URL
  form_url: "http://example.com/form"    # 登录后要访问的表单页面URL
  username: "your_username"  # 登录用户名
  password: "your_password"  # 登录密码

  # 登录表单元素定位器（根据实际登录页面修改）
  elements:
    username_field:
      locator: "id"       # 用户名输入框的定位方式
      value: "username"   # 用户名输入框的定位器值

    password_field:
      locator: "id"
      value: "password"   # 密码输入框的定位器值

    login_button:
      locator: "css_selector"
      value: "button[type='submit']"  # 登录按钮的定位器

  # 登录成功验证（用于判断是否登录成功）
  success_indicator:
    type: "url_contains"  # 验证方式：url_contains 或 element_exists
    value: "dashboard"    # URL中应包含的字符串（如登录后跳转到dashboard页面）
```

**如何获取登录表单元素定位信息？**

1. 打开登录页面
2. 右键点击用户名输入框 -> "检查"
3. 在开发者工具中查看元素的 `id`、`name` 等属性
4. 对密码框和登录按钮重复上述步骤

**登录成功验证说明：**
- `url_contains`：检查登录后URL是否包含特定字符串（如 "dashboard"、"home" 等）
- `element_exists`：检查登录后页面是否存在特定元素

#### 1.2 网站配置

```yaml
website:
  url: "https://example.com/form"  # 改为你的目标表单页面 URL
  name: "示例表单网站"
```

#### 1.3 表单元素配置

根据你的目标网页表单，配置表单元素的定位信息：

```yaml
form_elements:
  # 字段名: 对应 Excel 中的列名
  name:
    type: "input"           # 元素类型: input, select, textarea, button
    locator: "id"          # 定位方式: id, name, xpath, css_selector, class_name
    value: "name"          # 定位器的值（如 id="name"）

  email:
    type: "input"
    locator: "id"
    value: "email"

  # 下拉框示例
  gender:
    type: "select"
    locator: "id"
    value: "gender"

  # 提交按钮
  submit_button:
    type: "button"
    locator: "css_selector"
    value: "button[type='submit']"
```

**如何获取元素定位信息？**

1. 打开目标网页
2. 右键点击表单元素 -> "检查"
3. 在开发者工具中查看元素的 `id`、`name` 等属性
4. 推荐优先使用 `id` 定位，其次是 `name`、`css_selector`

#### 1.4 浏览器配置

```yaml
browser:
  headless: false        # true=后台运行不显示浏览器，false=显示浏览器窗口
  window_size: "1920,1080"
  timeout: 30           # 元素查找超时时间（秒）
```

### 2. 准备输入数据

在 `data/input_data.xlsx` 中准备要填写的数据：

| name | email | phone | gender | message |
|------|-------|-------|--------|---------|
| 张三 | zhangsan@example.com | 13800138000 | 男 | 测试消息1 |
| 李四 | lisi@example.com | 13900139000 | 女 | 测试消息2 |

**注意：** Excel 的列名必须与 `config.yaml` 中 `form_elements` 的字段名一致。

## 使用方法

### 运行程序

```bash
python main.py
```

### 运行流程

1. 程序读取配置文件
2. 读取 Excel 数据文件
3. 启动 Chrome 浏览器
4. **执行登录流程（如果启用）**
   - 打开登录页面
   - 填写用户名和密码
   - 点击登录按钮
   - 验证登录是否成功
   - 导航到表单页面
5. 逐行填写表单并提交
6. 将处理结果导出到 `output/results_*.xlsx`
7. 关闭浏览器

### 查看结果

处理完成后，在 `output/` 目录查看结果文件，包含以下信息：

- 原始数据
- 处理状态（成功/失败）
- 处理消息
- 处理时间

**快速打开结果：**
```bash
# 自动打开最新结果文件
python open_result.py

# 列出所有结果文件
python open_result.py --list
```

**自动打开功能：**
- 程序执行完成时，对话框会显示结果文件路径
- 点击"我已完成，继续"后，结果文件夹会自动打开
- Windows系统会自动选中结果文件

### 查看日志

详细的运行日志保存在 `logs/app.log`，可以查看每一步的执行情况和错误信息。

## 使用示例

### 示例 1：填写联系表单

假设目标网页有以下表单：

```html
<form>
  <input id="name" type="text" />
  <input id="email" type="email" />
  <textarea id="message"></textarea>
  <button type="submit">提交</button>
</form>
```

**配置文件 `config/config.yaml`：**

```yaml
website:
  url: "https://example.com/contact"

form_elements:
  name:
    type: "input"
    locator: "id"
    value: "name"

  email:
    type: "input"
    locator: "id"
    value: "email"

  message:
    type: "textarea"
    locator: "id"
    value: "message"

  submit_button:
    type: "button"
    locator: "css_selector"
    value: "button[type='submit']"
```

**数据文件 `data/input_data.xlsx`：**

| name | email | message |
|------|-------|---------|
| 张三 | zhangsan@example.com | 你好，这是测试消息 |
| 李四 | lisi@example.com | 请联系我 |

运行 `python main.py` 即可自动批量填写提交。

## 常见问题

### 1. 找不到 ChromeDriver

**错误：** `FileNotFoundError: 找不到 ChromeDriver`

**解决：**
- 确保已下载 ChromeDriver 并放到 `chromedriver-win64/` 目录
- 确保 ChromeDriver 版本与 Chrome 浏览器版本匹配

### 2. 元素定位失败

**错误：** `TimeoutException: 超时：找不到元素`

**解决：**
- 检查配置文件中的定位器是否正确
- 使用浏览器开发者工具确认元素属性
- 尝试增加 `timeout` 配置值
- 尝试使用其他定位方式（如从 `id` 改为 `css_selector`）

### 3. Excel 文件读取失败

**错误：** `FileNotFoundError: 数据文件不存在`

**解决：**
- 确保 `data/input_data.xlsx` 文件存在
- 检查配置文件中的文件路径是否正确

### 4. 编码问题

**错误：** 中文乱码

**解决：**
- 确保所有文件使用 UTF-8 编码
- 已在代码中配置 UTF-8 编码处理

### 5. 登录失败

**错误：** `登录失败，程序终止`

**解决：**
- 检查用户名和密码是否正确
- 检查登录表单元素定位器是否正确
- 检查登录成功验证配置（`success_indicator`）是否正确
- 使用浏览器开发者工具查看登录后的URL变化或页面元素
- 临时将 `headless` 设置为 `false`，观察登录过程

### 6. 如何处理验证码？

目前不支持自动识别验证码，需要：
- 方案1：联系网站管理员，获取测试环境（无验证码）
- 方案2：手动填写验证码后暂停程序
- 方案3：集成第三方验证码识别服务（需自行开发）

## 进阶配置

### 使用 XPath 定位元素

如果元素没有 `id` 或 `name` 属性，可以使用 XPath：

```yaml
form_elements:
  name:
    type: "input"
    locator: "xpath"
    value: "//input[@placeholder='请输入姓名']"
```

### 使用 CSS Selector 定位

```yaml
form_elements:
  submit_button:
    type: "button"
    locator: "css_selector"
    value: "#form-container > button.btn-submit"
```

### 无头模式运行

不显示浏览器窗口，后台运行：

```yaml
browser:
  headless: true
```

## 开发与扩展

### 添加新功能

可以在 `src/` 目录下添加新模块，例如：
- 验证码识别模块
- 截图功能
- 错误重试机制
- 多线程并发处理

### 支持其他浏览器

修改 `src/driver_manager.py`，添加 Firefox、Edge 等浏览器支持。

## 技术栈

- Python 3.x
- Selenium 4.x
- Pandas（数据处理）
- PyYAML（配置文件解析）
- openpyxl（Excel 读写）

## 许可证

本项目仅供学习和合法用途使用，请勿用于非法爬虫或攻击行为。

## 更新日志

### v1.0.0 (2025-10-30)
- 初始版本发布
- 支持 Excel/CSV 数据读取
- 支持基本表单元素填写
- 支持结果导出
