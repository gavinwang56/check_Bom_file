import pandas as pd
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# 隐藏主窗口
Tk().withdraw()

# 弹出文件选择对话框
file_path = askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xls;*.xlsx")])
if not file_path:
    print("未选择文件，程序结束。")
    exit()

# 读取Excel表格
df = pd.read_excel(file_path, header=None)

# 从第6行开始读取数据
data = df.iloc[5:]  # iloc中索引从0开始，所以5代表第6行

# 遍历数据行
for index, row in data.iterrows():
    if len(row) < 6:  # 确保该行至少有6列
        continue

    col3_value = str(row[2]).strip()  # 第3列的数据，转为字符串并去掉两端空格
    col4_value = str(row[3]) if pd.notna(row[3]) else ""  # 第4列数据，确保处理为字符串
    col5_value = str(row[4]) if pd.notna(row[4]) else ""  # 第5列数据，确保处理为字符串
    expected_count = row[5]    # 第6列的数据，作为期望数量

    # 检查第4列是否为空
    if col4_value:
        # 提取第4列的数字（正则表达式查找数字）
        numbers = re.findall(r'\d+', col4_value)

        # 检查提取的数字是否在当前行的第3列中
        for number in numbers:
            if number not in col3_value:
                print(f"错误: 在第{index + 1}行物料规格写的 '{col3_value}' 和封装 '{number}' 不匹配!")

    # 初始化总计数
    total_count = 0
    
    # 检查第5列是否为空
    if col5_value:
        items = col5_value.split()  # 分割第5列的值

        for item in items:
            if '-' in item:  # 处理范围，例如 "LED341-346"
                start, end = map(int, re.findall(r'\d+', item))
                total_count += (end - start + 1)  # 计算范围内的个数
            else:  # 处理单个项，例如 "R216"
                total_count += 1

        # 比较计算的数量与第6列的值
        if total_count != expected_count:
            print(f"错误: 在第{index + 1}行位号算的数量是 {total_count} 与写的数量 {expected_count} 不匹配!")

    # 判断第3列中的逗号后的字符串是否在逗号前面的字符串中
    if ',' in col3_value:
        before_comma, after_comma = col3_value.split(',', 1)
        after_comma = after_comma.strip()  # 去掉逗号后两端空格

        # 检查after_comma是否为数字加字母
        if re.match(r'^\d+[A-Za-z]', after_comma):
            # 判断逗号后面的字符串是否在逗号前面的字符串中，忽略大小写
            if after_comma.lower() not in before_comma.lower():
                print(f"错误: 在第{index + 1}行的物料规格写的 '{after_comma}' 和 '{before_comma}'参数不同!")

# 结束
print("检查完成。")

# 等待用户按下回车键后退出
input("按回车键退出...")
