import pandas as pd

# 读取两个 Excel 文件
excel1 = pd.read_excel("")  # 这是要更新的文件（目标文件）
excel2 = pd.read_excel("")  # 这是提供标题和日期的文件（源文件）

# 假设两个文件中都有 '链接', '标题', '日期' 列
# 遍历 excel1 中的每一行，根据 '链接' 查找 excel2 中对应的 '标题' 和 '日期'，并更新

# 使用 '链接' 列进行更新
for index, row in excel1.iterrows():
    link = row['链接']  # 获取目标文件中的链接
    # 在源文件中找到对应链接的标题和日期
    matching_row = excel2[excel2['链接'] == link]

    if not matching_row.empty:  # 如果找到匹配的链接
        # 更新目标文件中的 '标题' 和 '日期' 列
        excel1.at[index, '标题'] = matching_row.iloc[0]['标题']
        excel1.at[index, '日期'] = matching_row.iloc[0]['日期']

# 保存更新后的 DataFrame 到新 Excel 文件
excel1.to_excel('updated_file1.xlsx', index=False)

print("数据已成功更新并保存到 'updated_file1.xlsx'.")
