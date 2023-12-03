# TemplateCode


该库包含了各种类型的模板代码，供写程序时CTRL+C CTRL+V ，修改后使用

## 文件夹说明

### tkinter界面模板
**现成的图形化界面，可复制修改相关数据后使用**

`tkinter基本界面.py` : 一个基本的tkinter图形界面

`tkinter选择文件版界面.py` ：一个包含了选择文件框的tkinter图形界面


### pysimplegui界面模板
**使用pysimplegui库创建图形化界面，更为简便**

`pysimplegui界面.py`：一个包含了选择文件框等基本元素的pysimplegui图形界面

### 套打
**套打模板，用于批量套打文件**

`excel套打至word.py`:将一个excel数据表格多行的内容分别套打到word文档中，excel数据表格每行生成一份套打的word文档

`excel套打至excel.py`: 将一个excel数据表格多行的内容分别套打到excel 模板中，excel数据表格每行生成一份套打的excel模板文档

### 文件操作
**包括常用的文件操作函数如下：**

`get_allfolder_fullfilename(folder_path: str, filetype: list) -> list:`

获取待处理文件夹下面所有文件，并返回文件夹里面指定类型的文件名列表（包含子文件夹里面的文件）

---
`get_singlefolder_fullfilename(folder_path: str, filetype: list) ->list:`

获取待处理文件夹里指定后缀的文件名（单个文件夹，不包括子文件夹的文件）

---
`get_normal_filename(fullfilename: str) -> str:`

 从全文件名（包含绝对路径的文件名）转换为普通文件名,如："K:/Project/FilterDriver/DriverCodes/hello.txt" 通过转换变成"hello.txt"

---
`record_csv(content_list: list, csv_filename: str) -> None:`

将content_list列表的内容按行 追加 写入csv文件中(开始时默认保留第一行标题行，从第二行开始写入）。一般为一维列表 ，如['a','b','c','d','e','f']，将上述abcdef字母填进csv中的ABCDEF列中；如果用了二维列表，如[['a','b','c'],['d','e','f']]，则对应行A列单元格内容为字符串['a','b','c']，B列单元格内容为字符串['d','e','f']。

csv文件编码：utf-8

---
`record_txt(content: str, txt_filename: str) -> None:`

将content字符串内容按行追加写入txt文件中。txt文件编码：utf-8





