from os import listdir
from numpy import array
from numpy import loadtxt
from pandas import read_excel
from re import split
import tkinter as tk
from tkinter.filedialog import askdirectory


def getFileName(totalFileDir, out2):  # 'D:\\网课考勤记录\\tmp'
    attenFileName = '考勤表'  #
    listFile = listdir(totalFileDir)  # os.listdir：输出指定文件路径下的所有文件、文件夹，type：list
    xlsxList = []
    leaveTxt = []
    for file in listFile:
        if attenFileName in file:  # 此判断是为了找出考勤表，注意：考勤表 ！= 出勤表
            attenFileName = totalFileDir + '\\' + file
            out2.delete(0.0, tk.END)  # TK:清除原内容
            out2.insert('end', totalFileDir + '\\' + file + '考勤表读取成功！')
        elif '.xlsx' in file and '考勤表' not in file:
            xlsxList.append(totalFileDir + '\\' + file)  # 将所有xlsx放到列表容器中
            out2.insert('end', totalFileDir + '\\' + file + '出勤表读取成功！')
        else:
            out2.insert('end', '%s不是受支持的".csv"或者".xlsx"格式!' % (file))
    return attenFileName, xlsxList, leaveTxt  # 以后开放csvList接口，支持csv格式


def getStudName(attenFileName, xlsxList, leaveTxt, out2):
    out2.delete(0.0, tk.END)  # TK:清除内容

    # 读取出勤表获取学生姓名
    stuTable = read_excel(attenFileName, index_col=0)
    dataFroam = stuTable.iloc[2:, 1]
    stuNameList = list(array(dataFroam))  # 获取到学生列表

    # 请假名单读取：
    loadLeave = str(loadtxt(leaveTxt, encoding='utf-8', dtype='str', delimiter=','))
    todayLeave = split(r'2020/5/9(.*?)', loadLeave)[-1]

    outDict = {}  # 用于存储{表名：旷课名字}

    for xlsx in xlsxList:
        className = xlsx.split('\\')[-1]
        tmpNameList = stuNameList.copy()

        teachTable = read_excel(xlsx)
        arrTable = array(teachTable).reshape(1, -1)
        strTable = str(arrTable)
        splitTable = split(r'(?:-|－|\s|\+|\d|_)\s*', strTable)

        for name in stuNameList:
            if any(name in s for s in splitTable):
                tmpNameList.remove(name)
                outDict[className] = tmpNameList

    out2.insert('end', outDict, todayLeave)


# 实例化
root = tk.Tk()
root.title('易考勤    V 1.0.0')
root.geometry('800x500')  # 设置大小标题

lab1 = tk.Label(root, text='请选择考勤路径：', font=('宋体', 12), width=15, height=2)  # 组件
lab1.place(x=3, y=2)  # 放置标签,


def selectPath():
    # 选择文件path_接收文件地址
    path_ = askdirectory()
    # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
    path_ = path_.replace("/", "\\\\")
    # path设置path_的值
    path.set(path_)


path = tk.StringVar()
ent1 = tk.Entry(root, textvariable=path, bd=3, xscrollcommand=100, font=('宋体', 15), width=57)  # 获取用户输入组件到框体
ent1.place(x=125, y=11)

but1 = tk.Button(root, text='选择路径', font=('宋体', 12), width=10, height=1, command=selectPath)  # 点击组件
but1.place(x=700, y=11)

out2 = tk.Text(root, width=86, height=22, font=('宋体', 13))
out2.place(x=10, y=102)


def onClick1():  # 用于but1的点击事件
    s, k, l = getFileName(ent1.get(), out2)
    return s, k, l


but1 = tk.Button(root, text='读取数据', font=('宋体', 20), command=onClick1, width=10, height=1)
but1.place(x=220, y=40)


def onClick2():  # 用于but2的点击事件
    s, k, l = getFileName(ent1.get(), out2)
    getStudName(s, k, l, out2)


but2 = tk.Button(root, text='开始执行', font=('宋体', 20), command=onClick2, width=10, height=1)
but2.place(x=420, y=40)

root.mainloop()