# 数据文件说明

## 创建输入数据文件

### 方法 1：使用提供的 CSV 示例

1. 打开 `input_data_example.csv`
2. 修改数据内容
3. 另存为 `input_data.xlsx`（Excel 格式）

### 方法 2：直接创建 Excel 文件

1. 打开 Excel
2. 创建表格，第一行为列名（必须与配置文件中的字段名一致）
3. 从第二行开始填写数据
4. 保存为 `input_data.xlsx`

### 示例表格结构

| name | email | phone | gender | message |
|------|-------|-------|--------|---------|
| 张三 | zhangsan@example.com | 13800138000 | 男 | 测试消息1 |
| 李四 | lisi@example.com | 13900139000 | 女 | 测试消息2 |
| 王五 | wangwu@example.com | 13700137000 | 男 | 测试消息3 |

### 注意事项

1. **列名必须匹配**：Excel 的列名必须与 `config/config.yaml` 中 `form_elements` 的字段名完全一致
2. **不要有空行**：数据应该连续，避免中间有空行
3. **数据类型**：文本、数字都可以，程序会自动转换为字符串
4. **文件名**：默认为 `input_data.xlsx`，如需修改请同步修改配置文件

## 使用 Python 生成示例 Excel 文件

如果已安装 Python 环境，可以运行以下命令快速生成示例文件：

```python
import pandas as pd

# 创建示例数据
data = {
    'name': ['张三', '李四', '王五'],
    'email': ['zhangsan@example.com', 'lisi@example.com', 'wangwu@example.com'],
    'phone': ['13800138000', '13900139000', '13700137000'],
    'gender': ['男', '女', '男'],
    'message': ['测试消息1', '测试消息2', '测试消息3']
}

df = pd.DataFrame(data)
df.to_excel('input_data.xlsx', index=False, engine='openpyxl')
print('示例数据文件创建成功！')
```

将上述代码保存为 `create_sample.py`，在 data 目录下运行：

```bash
python create_sample.py
```
