# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
from xlsxwriter.workbook import Workbook
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.




column_data = []
workbook = Workbook("这是修改后的列需要加上列标题后手动复制回去.xlsx")
worksheet = workbook.add_worksheet()
red = workbook.add_format({'color': 'red'})



def readExcelColumn():
    global column_data
    global worksheet

    # 获取某一列的数据并转化为列表
    excelPath =  input("请输入文件路径+文件名 如：C: \\Users\LKH\excel名字.xlsx\n")
    columnStr = input("请输入要修改的列名，如 标题\n")
    excel = pd.read_excel(excelPath)
    column_data = excel[columnStr].tolist()
    print("reading this column")
    print(f"列行数总数：{len(column_data)}")
    print("reading done")
    # print(column_data)

def changeColor():
    global column_data
    changeStr = input("请输入要改为红色的文字\n")
    # global changeStr = '金'
    global workbook
    global worksheet

    columnDataStr = [str(i) for i in column_data]


    # print(columnDataStr)
    for row_num, sequence in enumerate(columnDataStr):

        format_pairs = []
        if(sequence.count(changeStr)==0):
           for base in sequence:
               if(base!=''):
                    format_pairs.append(base)
        else:
        # Get each DNA base character from the sequence.
        #去空串
            se = sequence.replace(changeStr,"*")
            list1 = []
            for base in se:
                if (base != ''):
                    list1.append(base)

            for i in list1:
                if i =="*":
                    format_pairs.extend((red, changeStr))
            # for base in sequence.split(changeStr):
            #     if base == '':
            #         format_pairs.extend((red, changeStr))
                else:
                     format_pairs.append(i)
                # Prefix each base with a format.
        # l = list(filter(None,format_pairs))
        worksheet.write_rich_string(row_num, 0, *format_pairs)
    workbook.close()
    print("修改成功")



    # Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('Bella is the best!')
    readExcelColumn()
    changeColor()


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
