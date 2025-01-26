import pandas as pd

# 读取 Excel 文件
df = pd.read_excel('your_file')


duplicates= df[df.duplicated(subset=[col for col in df.columns if col not in ['链接', '标题','日期',"发言人",'媒体']], keep=False)]
duplicates.to_excel("your-file", index=False)
df_cleaned = df.drop_duplicates(subset=[col for col in df.columns if col not in ['链接', '标题','日期',"发言人",'媒体']])

# 如果你想在去重后保存一个新的文件
df_cleaned.to_excel("new-file", index=False)

print("ok")
