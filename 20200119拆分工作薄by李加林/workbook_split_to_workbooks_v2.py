


import pandas as pd
import tkinter as tk
import tkinter.filedialog

class WorkbookSplit(object):
    """
    将单个工作薄按列字段拆分为多个工作薄
    """

    def __init__(self):
        """
        创建界面
        """
        # 新建窗口
        self.master = tk.Tk()
        # self.path 用于存放选择的文件路径
        # self.col 用于用户输入列字段
        # 注意！这些都是tk.StringVar()对象，不是str，其他地方要用的话要用get()方法获取str
        self.path = tk.StringVar()
        self.flag = tk.StringVar()
        self.col = tk.StringVar()
        # 主界面布局
        self.main()
        # 框架布局
        self.frame()
        # 保持运行
        self.master.mainloop()


    def main(self):
        """
        主界面布局设置
        :return:
        :rtype:
        """
        self.master.geometry("800x600+600+100")
        self.master.title("拆分excel工作薄工具by李加林 gui版v2.0")
        # 图片贴上去
        self.photo = tk.PhotoImage(file="img\\xzpq.gif")
        imgLable = tk.Label(self.master, image=self.photo)
        imgLable.pack()
        # 把标题贴上去
        tk.Label(self.master,text="拆分xlsx单个工作薄为多个工作薄\n表头默认为1行\n\n",font=("黑体",20)).pack()


    def frame(self):
        """
        框架布局设置
        :return:
        :rtype:
        """
        # 框架贴上去，再在框架里添加Lable，Entry，Button等控件
        frame1 = tk.Frame(self.master)
        frame1.pack()
        # 输入框，标记，按键
        tk.Label(frame1, text="目标路径:", font=("黑体", 16)).grid(row=1, column=0)
        tk.Entry(frame1, textvariable=self.path, width=50).grid(row=1, column=1)
        tk.Button(frame1, text="路径选择", command=self.select_path, font=("黑体", 16)).grid(row=1, column=2,padx=10)
        tk.Button(frame1, text="开始处理", command=self.split, font=("黑体", 16)).grid(row=5, column=1)
        tk.Label(frame1, text="列字段：", font=("黑体", 16)).grid(row=2, column=0)
        tk.Entry(frame1, textvariable=self.col, width=20).grid(row=2, column=1)


    def split(self):
        """
        使用pandas模块对xlsx工作薄进行处理
        :return: None
        :rtype: None
        """

        wb = pd.read_excel(self.path.get())  #打开工作薄
        col = wb[self.col.get()].unique() #对单位列去重复
        for x in col:
            child_wb = wb[wb[self.col.get()] == x]    #循环，得到每一个单位列表，
            child_wb.to_excel(x+'.xlsx',index=False)    #将得到的表保存成Excel格式

    def select_path(self):
        """
        选择文件，并获取文件的绝对路径
        :return:
        :rtype:
        """
        # 选择文件，path_select变量接收文件地址
        path_select = tkinter.filedialog.askopenfilename()
        # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
        # 注意：\\转义后为\，所以\\\\转义后为\\
        path_select = path_select.replace("/", "\\\\")
        # self.path设置path_select的值
        self.path.set(path_select)

if __name__ == '__main__':
    # 实例化运行
    app = WorkbookSplit()


































# import pandas as pd #导入模块
# wb = pd.read_excel(r'd:\Me_py\everyday_py\资料\2018-04-08.xlsx')  #打开工作薄
# col = wb['地市'].unique() #对地市列去重复
# new_wb = pd.ExcelWriter('拆分表.xlsx')      #新建一个工作薄
# for x in col:
#     child_wb = wb[wb['地市'] == x]    #循环，得到每一个地市列表，
#     child_wb.to_excel(new_wb,index=False,sheet_name=x)  #将得到的表保存成Excel格式
# new_wb.save()