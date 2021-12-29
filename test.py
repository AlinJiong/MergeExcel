import pandas as pd

df = pd.read_excel('data\口岸运营部劳务派遣人员名单（1）.xlsx', header=1)
df.head()

print(df.columns[1])
print(type(df.columns[1]))