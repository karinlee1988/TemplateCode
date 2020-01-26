import tkinter as tk
import tkinter.filedialog

class GuiTemp(object):
    """
    tkinter GUI界面模板（可选择文件进行操作）
    """

    def __init__(self):
        """
        创建界面
        """
        # 新建窗口
        self.master = tk.Tk()
        # self.path 用于存放选择的文件路径
        # self.flag暂时没用
        # self.v1 self.v2接收用户参数
        # 注意！这些都是tk.StringVar()对象，不是str，其他地方要用的话要用get()方法获取str
        self.path = tk.StringVar()
        self.flag = tk.StringVar()
        self.v1 = tk.StringVar()
        self.v2 = tk.StringVar()
        # 窗口图片横幅
        self.img_pack()
        # 主界面布局
        self.window()
        # 框架布局
        self.frame()
        # 保持运行
        self.master.mainloop()

    def img_pack(self):
        """
        窗口图片横幅
        :return:
        :rtype:
        """
        # 图片贴上去
        self.photo = tk.PhotoImage(file="xzpq.gif")
        imgLable = tk.Label(self.master, image=self.photo)
        imgLable.pack()

    def window(self):
        """
        主界面布局设置
        :return:
        :rtype:
        """
        self.master.geometry("800x600+600+100")
        self.master.title("xxxxxxby李加林v1.0")
        # 图片贴上去
        # self.photo = tk.PhotoImage(file="xzpq.gif")
        # imgLable = tk.Label(self.master, image=self.photo)
        # imgLable.pack()
        # 把标题贴上去
        tk.Label(self.master,text="------这个是大标题-------",font=("黑体",20)).pack()
        tk.Label(self.master,text="------这个是小标题-------",font=("黑体",16)).pack()

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
        tk.Button(frame1, text="路径选择", command=self.select_path, font=("黑体", 16)).grid(row=1, column=2)
        tk.Label(frame1, text="参数1", font=("黑体", 16)).grid(row=2, column=0)
        tk.Entry(frame1, textvariable=self.v1).grid(row=2, column=1)
        tk.Label(frame1, text="参数2:", font=("黑体", 16)).grid(row=3, column=0)
        tk.Entry(frame1, textvariable=self.v2).grid(row=3, column=1)
        # 按这个按钮执行主程序
        tk.Button(frame1, text="开始处理", command=self.main, font=("黑体", 16)).grid(row=4, column=1,pady=66)
        tk.Entry(frame1, textvariable=self.flag,state="readonly").grid(row=4, column=2,pady=66)

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

    def main(self):
        """
        这个是主程序
        :return:
        :rtype:
        """
        # 标志设置为处理完成
        self.flag.set("处理完成！")

if __name__ == '__main__':
    app = GuiTemp()
