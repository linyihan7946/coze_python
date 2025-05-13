import pandas as pd

# 读取 Excel 文件
excel_path = 'C:/Users/Administrator/Desktop/新建 XLSX 工作表.xlsx'
markdown_path = 'C:/Users/Administrator/Desktop/output.md'

sheet_name = 0
df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# 处理合并单元格
df = df.fillna(method='ffill', axis=0)

# 创建markdown内容
markdown_content = "# 目录结构\n\n"

# 用于跟踪当前目录
current_level1 = ""
current_level2 = ""
current_level3 = ""

# 遍历每一行
for index, row in df.iterrows():
    # 获取每一列的值
    level1 = str(row[0]).strip() if pd.notna(row[0]) else ""
    level2 = str(row[1]).strip() if pd.notna(row[1]) else ""
    level3 = str(row[2]).strip() if pd.notna(row[2]) else ""
    level4 = str(row[3]).strip() if pd.notna(row[3]) else ""
    
    # 只有当目录名称改变时才添加新的目录
    if level1 and level1 != current_level1:
        markdown_content += f"## {level1}\n"
        current_level1 = level1
        current_level2 = ""  # 重置二级目录
        current_level3 = ""  # 重置三级目录
    
    if level2 and level2 != current_level2:
        markdown_content += f"### {level2}\n"
        current_level2 = level2
        current_level3 = ""  # 重置三级目录
    
    # 合并第三列和第四列的内容
    combined_level3 = level3
    if level4:
        combined_level3 = f"{level3}：{level4}"
    
    if combined_level3 and combined_level3 != current_level3:
        markdown_content += f"#### {combined_level3}\n"
        current_level3 = combined_level3
    
    # 添加空行以增加可读性
    markdown_content += "\n"

# 保存为 .md 文件
with open(markdown_path, 'w', encoding='utf-8') as f:
    f.write(markdown_content)

print("转换完成，已保存为 output.md") 