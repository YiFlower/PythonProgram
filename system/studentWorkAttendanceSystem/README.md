## 一、实现过程
### （一）数据读取
**导入相关库**
```python
from os import listdir
from numpy import array
from numpy import loadtxt
from pandas import read_excel
from re import split
import tkinter as tk
from tkinter.filedialog import askdirectory
```
**读取文件**

```python
def getFileName(totalFileDir,out2):#'D:\\网课考勤记录\\tmp'
    attenFileName = '考勤表'  # 
    listFile = listdir(totalFileDir)  # os.listdir：输出指定文件路径下的所有文件、文件夹，type：list
    xlsxList = []
    csvList = []
    leaveTxt = []
    for file in listFile:
        if attenFileName in  file:  #此判断是为了找出考勤表，注意：考勤表 ！= 出勤表
            attenFileName = totalFileDir + '\\' + file
            out2.delete(0.0, tk.END)  # TK:清除原内容
            out2.insert('end',totalFileDir + '\\' + file+'考勤表读取成功！')
        elif '.xlsx' in file and '考勤表' not in file:
            xlsxList.append(totalFileDir + '\\' + file)  #将所有xlsx放到列表容器中
            out2.insert('end',totalFileDir + '\\' + file+'出勤表读取成功！')
        else:
            out2.insert('end','%s不是受支持的".csv"或者".xlsx"格式!'%(file))
    return attenFileName,xlsxList,leaveTxt  #以后开放csvList接口，支持csv格式
```

### （二）数据切片并对比匹配
**获取出勤表学生列，使用正则表达式切片。最开始想的是使用python自带的strip或者replace之类的内置函数，但是只支持单个字符，故使用正则表达式的split函数。**

```python
def getStudName(attenFileName,xlsxList,leaveTxt,out2):
    out2.delete(0.0, tk.END)  # TK:清除内容

    #读取出勤表获取学生姓名
    stuTable = read_excel(attenFileName,index_col=0)
    dataFroam = stuTable.iloc[2:,1]
    stuNameList = list(array(dataFroam)) #获取到学生列表

    #请假名单读取：
    loadLeave = str(loadtxt(leaveTxt,encoding='utf-8',dtype='str',delimiter=','))
    todayLeave = split(r'2020/5/9(.*?)', loadLeave)[-1]

    outDict = {}  # 用于存储{表名：旷课名字}

    for xlsx in xlsxList:
        className = xlsx.split('\\')[-1]
        tmpNameList = stuNameList.copy()

        teachTable = read_excel(xlsx)
        arrTable = array(teachTable).reshape(1,-1)
        strTable = str(arrTable)
        splitTable = split(r'(?:-|－|\s|\+|\d|_)\s*',strTable)

        for name in stuNameList:
            if any(name in s for s in splitTable):
                tmpNameList.remove(name)
                outDict[className] = tmpNameList

    out2.insert('end',outDict,todayLeave)
```

### （三）图形化开发
**实例化TK对象，设置基本属性**

```python
#实例化
root = tk.Tk()
root.title('易考勤    V 1.0.0')
root.geometry('800x500')  #设置大小标题

lab1 = tk.Label(root,text = '请选择考勤路径：',font = ('宋体',12),width = 15,height = 2)  #组件
lab1.place(x=3,y=2)  #放置标签,
```
**设置文件读取目录，可手动选择更改**

```python
def selectPath():
    # 选择文件path_接收文件地址
    path_ = askdirectory()
    # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
    path_ = path_.replace("/", "\\\\")
    # path设置path_的值
    path.set(path_)
```
**设置界面按钮大小等属性**
```python
path = tk.StringVar()
ent1 = tk.Entry(root,textvariable = path,bd = 3,xscrollcommand = 100,font=('宋体',15),width = 57)  #获取用户输入组件到框体
ent1.place(x=125,y=11)

but1 = tk.Button(root,text = '选择路径',font = ('宋体',12),width = 10,height = 1,command = selectPath)  #点击组件
but1.place(x=700,y=11)

out2 = tk.Text(root,width=86,height=22,font=('宋体',13))
out2.place(x=10,y=102)
```
**绑定点击事件，接收函数输出**
```python
def onClick1():#用于but1的点击事件
    s,k,l= getFileName(ent1.get(), out2)
    return s,k,l
but1 = tk.Button(root,text = '读取数据',font = ('宋体',20),command=onClick1,width = 10,height = 1)
but1.place(x=220,y=40)

def onClick2():#用于but2的点击事件
    s, k, l = getFileName(ent1.get(), out2)
    getStudName(s,k,l,out2)

but2 = tk.Button(root,text = '开始执行',font = ('宋体',20),command=onClick2,width = 10,height = 1)
but2.place(x=420,y=40)

root.mainloop()
```

## 二、运行效果
### （一）界面
![运行界面](https://img-blog.csdnimg.cn/20200630113832779.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQ0NDkxNzA5,size_16,color_FFFFFF,t_70#pic_center)

### （二）选择路径
**推荐创建专用目录，里面放置需要考勤表和出勤表，输出结果后移出。**
![选择路径](https://img-blog.csdnimg.cn/20200630114538258.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQ0NDkxNzA5,size_16,color_FFFFFF,t_70#pic_center)
### （三）读取数据
**选择路径后点击读取数据即可
![在这里插入图片描述](https://img-blog.csdnimg.cn/20200630114849440.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQ0NDkxNzA5,size_16,color_FFFFFF,t_70#pic_center)
### （四）最终效果
![考勤结果](https://img-blog.csdnimg.cn/20200630115036536.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQ0NDkxNzA5,size_16,color_FFFFFF,t_70#pic_center)
## 三、注意事项
1. 读取考勤表的关键字是”考勤表“，只要文件名包含这三个字即可；不过需要注意，除了考勤表，其他文件名不得包含这个关键字。

```python
attenFileName = '考勤表'
    for file in listFile:
        if attenFileName in  file:
```
2. 请假名单的功能是直接写死了的，有需要的网友可以使用其他方法。

```python
    loadLeave = str(loadtxt(leaveTxt,encoding='utf-8',dtype='str',delimiter=','))
    todayLeave = split(r'2020/5/9(.*?)', loadLeave)[-1]
```
3. 为了更加的方便，我把这个直接打包成了.exe（是真的大，300+MB），创建快捷方式后可直接双击运行（==需要安装pyinstaller库编译==)。![执行文件](https://img-blog.csdnimg.cn/20200630115858691.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQ0NDkxNzA5,size_16,color_FFFFFF,t_70#pic_center)
