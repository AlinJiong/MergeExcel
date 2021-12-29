import pandas as pd
import os


#使用os模块walk函数，搜索出某目录下的全部excel文件
def getFileName(filepath):
    file_list = []
    for root, dirs, files in os.walk(filepath):
        print("使用前，请把需要合并的excel文件放到data目录下面！")
        print('读取到的文件如下：')
        for filespath in files:
            print(os.path.join(root, filespath))
            file_list.append(os.path.join(root, filespath))

    print('*' * 100)
    return file_list


#合并excel
def MergeExcel(filepath):
    file_list = getFileName(filepath)

    merge = []
    #合并多个excel文件
    header = input(
        '先确认表格最后列是否有多余的备注信息，若有，请先删除，并关闭所有文件。\n请输入需要跳过的行，例如前两行输入2(默认为0， 即不跳过，按回车键执行):\n'
    ) or 0
    print('*' * 100)
    header = int(header)

    df = pd.read_excel(file_list[0], header=header)
    print(
        '尝试读取第一个文件，注意查看信息是否正确，是否有多余列！\n如果有，输入正确的起始列和结束列+1，第一列为0，第二列为1....\n例如:0 9表示前九列，2 9表示第三列到第九列'
    )

    print('*' * 100)
    print(df.columns)
    print('当前列长度为%s，若出现Unnamed,则需要调整起始列和结束列!!!!!' % (len(df.columns)))
    print('*' * 100)

    print('请输入正确的起始列和结束列，中间用空格键隔开，此处建议输入0 所需列数(若无问题，按回车键)：')
    start, end = input().split() or (0, len(df.columns))
    start, end = int(start), int(end)

    if (start, end) != (0, len(df.columns)):
        df = df.iloc[:, start:end]
        print('+' * 100)
        print('新的列如下：')
        print('*' * 100)
        print(df.columns)

    for each in file_list:
        #读取xlsx格式文件
        df = pd.read_excel(each, header=header)
        df = df.iloc[:, start:end]
        merge.append(df)

    writer = pd.ExcelWriter('合并.xlsx')
    data = pd.concat(merge)

    print('*' * 100)
    print('列数和列名对应如下：')
    count = 0
    for i in data.columns:
        print(count, i)
        count += 1
    drop_name = int(input('请输入去重的列:(输入身份证或者电话号码关键列对应的列数即可)\n'))

    data.drop_duplicates(subset=[data.columns[drop_name]], inplace=True)

    reset_name = int(input('请输入重新建立序号的列：(即序号、编号等)\n'))
    data[data.columns[reset_name]] = range(len(data))

    data.to_excel(writer, 'Sheet1', index=False)
    writer.save()
    print('操作成功！')


#调取方法，合并数据
filepath = 'data\\'
MergeExcel(filepath)