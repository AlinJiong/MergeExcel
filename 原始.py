import os, pandas as pd


# 获取文件夹下文件全路径名
def get_files(path):
    fs = []
    for root, dirs, files in os.walk(path):
        for file in files:
            fs.append(os.path.join(root, file))
    return fs


# 合并文件
def merge():
    files = get_files('D:/excels')
    arr = []
    for i in files:
        arr.append(pd.read_excel(i))
    writer = pd.ExcelWriter('D:/excels/merge.xlsx')
    pd.concat(arr).to_excel(writer, 'Sheet1', index=False)
    writer.save()


if __name__ == '__main__':
    merge()