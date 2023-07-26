import tkinter as tk
from tkinter.ttk import *
from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import os
import threading
import time


# 截取字符串。。原始，查找前字符串位置。。查找后字符串位置。。返回中间的内容
def strGetlen(strn, strx, strend):
    sint = strn.find(strx)
    strn = strn[sint + len(strx):]
    if strend == "":
        return strn
    eint = strn.find(strend)
    return strn[:eint]


# 将函数打包进线程
def thread_is(func):
    '''将函数打包进线程'''
    # 创建
    t = threading.Thread(target=func)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()
    # 阻塞--卡死界面！
    # t.join()


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.window_init()
        self.create_notebook()

    # 界面设置
    def window_init(self):
        self.master.title('工具')  # 标题
        width, height = self.master.maxsize()  # 获取窗体大小
        # self.width = int(width / 1.5)
        # self.height = int(height / 1.5)
        # self.master.geometry("{}x{}".format(self.width, self.height))  # 设置大小
        # self.master.resizable(False, False)  # 固定窗体
        # self.master.iconbitmap("C:\\Users\\Administrator\\Desktop\\tk_a16-62\\123.png")#图标
        # width, height=1360,1024
        if width < 1152:
            width, height = 1152, 864
        self.width = int(width / 2)
        self.height = int(height / 2)
        widths = int(self.width / 2)  # 屏幕显示位置
        heights = int(self.height / 2)  # 屏幕显示位置
        self.master.geometry("{}x{}+{}+{}".format(self.width, self.height, widths, heights))  # 设置大小
        # 绑定窗口变动事件
        self.master.bind('<Configure>', self.WindowResize)
        self.dataA = None
        self.dataB = None
        self.replace = '----'

    # 窗口尺寸调整处理函数
    def WindowResize(self, event):
        # global save_width
        # global save_height

        new_width = self.master.winfo_width()
        new_height = self.master.winfo_height()

        if new_width == 1 and new_height == 1:
            return
        if self.width != new_width or self.height != new_height:
            # self.master.config(width=new_width - 40, height=new_height - 60)
            # button1.place(x=20, y=new_height - 40)
            self.width = new_width
            self.height = new_height
            self.treeview.place(x=0, y=37, width=self.width, height=self.height / 2)
            self.TabStrip1.place(x=0, y=self.height / 1.8, width=self.width, height=self.height / 2.3)
            # self.create_notebook()

    # 笔记本组件类
    def create_notebook(self):
        but_open = Button(self.master, text='导入A', width=15, command=self.open_file)
        but_open.place(x=5, y=5)
        out_open = Button(self.master, text='导入B', width=15, command=self.open_files)
        out_open.place(x=115, y=5)
        # but_stat = Button(self.master, text='执行程序', width=15, command=self.open_file)
        # but_stat.place(x=230, y=5)

        # self.pb = Progressbar(self.master, length=200, mode="determinate", orient=HORIZONTAL)
        # self.pb.place(x=self.width-210, y=8)
        # # self.pb.place_forget()
        # self.label = Label(self.master)
        # self.label['text']="文件导入100/100"
        # self.label.place(x=self.width-320, y=8)

        columns = ("ID", "列1", "列2", "列3", "提示")
        self.treeview = Treeview(self.master, height=self.height, show="headings", columns=columns)  # 表格

        self.treeview.column("ID", anchor='center', width=10)  # 表示列,不显示
        self.treeview.column("列1", anchor='center', width=30)  # 表示列,不显示
        self.treeview.column("列2", anchor='center', width=30)
        self.treeview.column("列3", anchor='center', width=150)
        self.treeview.column("提示", anchor='center', width=150)
        #
        self.treeview.heading("ID", text="ID")  # 显示表头
        self.treeview.heading("列1", text="列1")
        self.treeview.heading("列2", text="列2")
        self.treeview.heading("列3", text="列3")
        self.treeview.heading("提示", text="提示")
        self.treeview.place(x=0, y=37, width=self.width, height=self.height / 2)

        # frame = Frame(self.master)
        # frame.place(x=0, y=self.height / 1.8, width=self.width, height=self.height / 2.3)
        # self.but_account = Button(frame, text='账号密码', width=15,
        #                           command=lambda: thread_is(self.out_but_account))
        # self.but_account.place(x=5, y=50)
        self.TabStrip1 = Notebook(self.master)
        self.TabStrip1.place(x=0, y=self.height / 1.8, width=self.width, height=self.height / 2.3)

        self.select_tab = Frame(self.TabStrip1)
        self.select_tab_label = Label(self.select_tab)
        self.select_tab_label.place(relx=0.1, rely=0.5)
        self.TabStrip1.add(self.select_tab, text='账号提取')
        self.but_account = Button(self.select_tab, text='执行', width=15,
                                  command=lambda: thread_is(self.out_but_account))
        self.but_account.place(x=5, y=100)

        # self.checkbutton_account_var = BooleanVar()
        # self.checkbutton_account = Checkbutton(self.select_tab, variable=self.checkbutton_account_var, text="账号密码")
        # self.checkbutton_account.place(x=5, y=5)

        self.radioText = StringVar()

        self.radiobutton1 = Radiobutton(self.select_tab, text="剔除列1", variable=self.radioText, value=0)
        self.radiobutton1.place(x=5, y=5)

        self.radiobutton2 = Radiobutton(self.select_tab, text="剔除列2", variable=self.radioText, value=1)
        self.radiobutton2.place(x=100, y=5)

        self.radiobutton3 = Radiobutton(self.select_tab, text="剔除列3", variable=self.radioText, value=2)
        self.radiobutton3.place(x=200, y=5)

        # self.radiobutton4 = Radiobutton(self.select_tab, text="账号密码", variable=self.radioText, value=3)
        # self.radiobutton4.place(x=300, y=5)
        self.updta_tab = Frame(self.TabStrip1)
        self.updta_tab_label = Label(self.updta_tab)
        self.updta_tab_label.place(relx=0.1, rely=0.5)
        self.TabStrip1.add(self.updta_tab, text='账号去重|替换')
        self.but_removal = Button(self.updta_tab, text='账号去重', width=15,
                                  command=lambda: thread_is(self.out_but_removal))
        self.but_removal.place(x=5, y=5)

        self.get_tab = Frame(self.TabStrip1)
        self.get_tab_label = Label(self.get_tab)
        self.get_tab_label.place(relx=0.1, rely=0.5)
        self.TabStrip1.add(self.get_tab, text='A去B')
        self.but_removal_a = Button(self.get_tab, text='A去B', width=15,
                                    command=lambda: thread_is(self.out_but_removal_a))
        self.but_removal_a.place(x=5, y=5)

    def progressbar(self):
        progress_bar = Progressbar(self.master, length=200, mode="determinate", orient=HORIZONTAL)
        # progress_bar.place(x=5, y=5)
        # progress_bar.place(x=px, y=py)
        progress_bar.place(x=235, y=8)

        progress_bar_label = Label(self.master)
        # progress_bar_label.place(x=210, y=5)
        # progress_bar_label.place(x=lx, y=ly)
        progress_bar_label.place(x=435, y=8)
        return progress_bar, progress_bar_label

    # 清空表中数据
    def delButton(self):
        x = self.treeview.get_children()
        for item in x:
            self.treeview.delete(item)

    def open_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[(" please open txt file", "*.txt")])
        if os.path.exists(self.file_path):
            self.dataA = self.open_write_text(self.file_path)

            # 判断是否选择正确文件
            if self.replace not in self.dataA[0]:
                tkinter.messagebox.showerror('提示', '请选择正确文件！')
            else:
                if os.path.exists(self.file_path):
                    self.delButton()
                start = time.perf_counter()
                scale = len(self.dataA)
                if scale > 100:
                    scale = 100
                progress_bar, progress_bar_label = self.progressbar()
                progress_bar["maximum"] = scale
                isCount = 1
                for content in self.dataA[:scale]:
                    content_list = content.replace('\r', '').replace('\n', '')
                    if content_list != '':
                        content_list = content_list.split(self.replace)
                        username, password, data, top = "", "", "", ""
                        if len(content_list) == 1:
                            username = content_list[0]
                        if len(content_list) == 2:
                            username = content_list[0]
                            password = content_list[1]
                        if len(content_list) == 3:
                            username = content_list[0]
                            password = content_list[1]
                            data = content_list[2]
                        if len(content_list) >= 4:
                            username = content_list[0]
                            password = content_list[1]
                            data = content_list[2]
                            top = content_list[3]
                        # data = self.replace.join(content_list[2:])
                        # 填入表格
                        self.treeview.insert('', isCount, values=(isCount, username, password, data, top))
                        t = time.perf_counter() - start
                        progress_bar["value"] = isCount  # 每次更新1
                        progress_bar_label['text'] = "文本导入:{}/{}".format(scale, isCount)
                        self.master.update()  # 更新画面
                        isCount += 1
                progress_bar.place_forget()
                progress_bar_label.place_forget()

    def open_write_text(self, filename, data=None):
        if data != None:
            mode = "w"
            with open(filename, mode, encoding="utf-8") as f:
                f.write(data)
            return
        else:
            mode = "r"
        try:
            with open(filename, mode) as f:
                data_file = f.read()

                # data_file = f.readlines()
        except:
            with open(filename, mode, encoding="utf-8") as f:
                data_file = f.read()
                # data_file = f.readlines()
        if data_file[-1:] == '\n':
            data_file = data_file[:-1]
        data_file = data_file.replace(" ", '').replace("\r", '').split('\n')
        return data_file

    def open_files(self):
        file_path = filedialog.askopenfilename(filetypes=[(" please open txt file", "*.txt")])
        if os.path.exists(file_path):
            self.dataB = self.open_write_text(file_path)

    # 导出数据
    def out_file(self):
        if self.treeview.selection() == ():
            tkinter.messagebox.showerror('提示', '请选择导出数据！')
        else:
            print(len(self.treeview.selection()))
            filename = tkinter.filedialog.asksaveasfilename(
                filetypes=[(" please open txt file", "*.txt")]) + '.txt'
            with open(filename, 'w', encoding='utf-8') as f:
                for item in self.treeview.selection():
                    save_ = self.treeview.item(item, "values")
                    if filename != '.txt':
                        f.write(save_[0] + self.replace + save_[1] + self.replace + save_[2] + self.replace + save_[
                            4] + '\n')
            tkinter.messagebox.showinfo('提示', '导出成功！')

    # 账号密码
    def out_but_account(self):
        if self.dataA == None:
            tkinter.messagebox.showerror('提示', '请导入数据A！')
        else:
            # 账号密码
            if self.radioText.get() != "":
                radioText = int(self.radioText.get())
            else:
                radioText = 999
            if radioText == 0:
                data = [self.replace.join(item.replace("\n", "").split(self.replace)[1:]) + "\n" for item in self.dataA]
                radioName = "剔除列1"
            elif radioText == 1:
                data = [self.replace.join(item.replace("\n", "").split(self.replace)[0::2]) + "\n" for item in
                        self.dataA]
                radioName = "剔除列2"
                for item in self.dataA:
                    txt = self.replace.join(item.replace("\n", "").split(self.replace)[0::2]) + "\n"
            elif radioText == 2:
                radioName = "剔除列3"
                data = [self.replace.join(item.replace("\n", "").split(self.replace)[:2]) + "\n" for item in self.dataA]
            else:
                data = [self.replace.join(item.replace("\n", "").split(self.replace)[:2]) + "\n" for item in self.dataA]
                radioName = "剔除列3"
            if data != "":
                filename = "{}[{}].txt".format(self.file_path.replace('.txt', ''), radioName)
                self.open_write_text(filename, ''.join(data))
                tkinter.messagebox.showinfo('提示', '导出完成！')

    def out_but_removal(self):
        from collections import Counter
        data_count = Counter(self.dataA)
        removar_data = [key for key, values in dict(data_count).items() if values > 1]
        out_removar_data = list(set(self.dataA))
        filename = "{}[已去重].txt".format(self.file_path.replace('.txt', ''))
        self.open_write_text(filename, "\n".join(out_removar_data))
        if removar_data != []:
            filename = "{}[重复项].txt".format(self.file_path.replace('.txt', ''))
            self.open_write_text(filename, "\n".join(removar_data))
        tkinter.messagebox.showinfo('提示', '导出完成！')
        # progress_bar.place_forget()
        # progress_bar_label.place_forget()

    def out_but_removal_a(self):
        if self.dataA == None or self.dataA == []:
            tkinter.messagebox.showerror('提示', '请导入数据A！')
        elif self.dataB == None or self.dataA == []:
            tkinter.messagebox.showerror('提示', '请导入数据B！')
        else:
            set_c = set(self.dataA) & set(self.dataB)
            data = list(set(self.dataA) - set_c)
            filename = "{}[A去B].txt".format(self.file_path.replace('.txt', ''))
            self.open_write_text(filename, "\n".join(data))
            tkinter.messagebox.showinfo('提示', '导出完成！')

    def create_widgets(self):
        pass


root = tk.Tk()
app = Application(master=root)
app.mainloop()
