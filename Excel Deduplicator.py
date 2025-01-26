import pandas as pd

# 读取 Excel 文件
df = pd.read_excel(r"C:\Users\86159\Desktop\外交语料\外交部\缺失链接√.xlsx")


duplicates= df[df.duplicated(subset=[col for col in df.columns if col not in ['链接', '标题','日期',"发言人",'媒体']], keep=False)]
duplicates.to_excel(r"C:\Users\86159\Desktop\外交语料\外交部\缺失链接√×.xlsx", index=False)
df_cleaned = df.drop_duplicates(subset=[col for col in df.columns if col not in ['链接', '标题','日期',"发言人",'媒体']])

# 如果你想在去重后保存一个新的文件
df_cleaned.to_excel(r"C:\Users\86159\Desktop\外交语料\外交部\缺失链接√√02.xlsx", index=False)

print("ok")
