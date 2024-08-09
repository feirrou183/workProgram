import easygui
import numpy as np
import openpyxl as xls
import pandas as pd
import copy
import os

pd.options.mode.chained_assignment = None


def openfile(excel):  # 返回值：excel表单对象
    # print(excel.sheet_names())
    sheet = excel.active
    return sheet


def create_excel_list(sheet):  # excel表单对象  返回值：数据列表
    cool_excel = pd.read_excel(sheet, sheet_name=0, engine='openpyxl')
    # 下面打印出第2列全部内容

    # 此时已经获得了完整的excel表格 cool_excel[][]获得了完整的表格。下面进行数据操作
    return cool_excel


def process_excel(cool_excel, N):  # cool_excel:导入的数据列表。 N：所需的月份数。  返回值：处理完成的新数据列表
    import copy
    cool_excel.fillna(0, inplace=True)
    '''
    实现功能的方法： 创建一个新列表cool_excel_1[][] 将处理完的数据放入
    '''

    cool_excel_1 = copy.deepcopy(cool_excel)
    cool_excel_1.iloc[:, 1:34] = 0

    '''
    到这里已经完成cool_excel_1的全部初始化工作（客户，备注等信息导入）。下面进行数据操作
    '''
    mid_data = 0.00  # mid_data是累积计算的存放器
    mid_data_1 = 0.00  # mid_data_1 是防止产生负数的保存值

    for i in range(cool_excel.shape[0]):  # 这里开始只对客户数据进行处理，故从第三行开始
        mid_data = cool_excel.iloc[i, 33]

        for j in range(N + 1):
            cool_excel_1.iloc[i, 31 - 2 * j] = cool_excel.iloc[i, 31 - 2 * j]

            mid_data_1 = mid_data  # 这里给mid_data_1赋值保证为前一个值
            if cool_excel_1.iloc[i, 31 - 2 * j] == np.nan:  # 空格为未填数据
                pass
            elif cool_excel_1.iloc[i, 31 - 2 * j] != np.nan:
                mid_data -= cool_excel_1.iloc[i, 31 - 2 * j]

            if mid_data <= 5 or j == N - 1:  # mid_data<= 5是给个最小值    j== N-1 是指已经满N个月
                if (mid_data > 0) and (mid_data <= 5):
                    break
                cool_excel_1.iloc[i, 31 - 2 * j] = mid_data_1
                break

        for k in range(N):  # 累积和
            if cool_excel_1.iloc[i, 31 - 2 * k] == np.nan:
                pass
            else:
                cool_excel_1.iloc[i, 33] += cool_excel_1.iloc[i, 31 - 2 * k]
    # 这里已经完成了对数据的处理操作，下面将cool_excel_1的全部数据导入一个新表中
    return cool_excel_1


def new_file():
    easygui.msgbox(msg='文件处理完成，请选择一个新的路径（推荐桌面）和填写新文件名')
    file_new_path = easygui.filesavebox('请选择文件保存路径')
    return file_new_path


def pressSheetAndSqueeze(cool_excel_1):
    df = cool_excel_1
    df_list = []

    # 剔除无关选项
    drop_column = []
    for j in range(1, 33):
        if (sum(df.iloc[:, j]) == 0):
            drop_column.append(j)

    df.drop(df.columns[drop_column], axis=1, inplace=True)
    df_list.append(df)

    class_list = df['跟单员'].value_counts().index.values.tolist()

    for person in class_list:
        temp_df = df[df["跟单员"] == person]
        df_list.append(temp_df)

    class_list.insert(0, "全表")
    return class_list, df_list


def main(month):  # N为需要保留的月份数量

    Work_Path = r"E:\work"
    os.chdir(Work_Path)
    # file = easygui.fileopenbox('')

    file = r'E:\work\2024\month07\day04\7月应收.xlsx'
    excel = xls.load_workbook(file, data_only=True)

    # sheet = openfile(excel)          #打开文件，得到表单对象
    cool_excel = create_excel_list(file)
    N = 18 - month
    # 下面实现功能
    cool_excel_1 = process_excel(cool_excel, N)  # 处理完成的cool_excel_1 返回
    class_list, df_list = pressSheetAndSqueeze(cool_excel_1)

    file_new_path = new_file()

    writer = pd.ExcelWriter(file_new_path)
    for i in range(len(class_list)):
        df_list[i].to_excel(writer, class_list[i])

    writer._save()
    writer.close()


month = 6  # 截止到几月的数据

main(month)
