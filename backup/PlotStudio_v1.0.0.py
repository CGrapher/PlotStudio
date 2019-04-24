import tkinter as tk
import xlwings as xw
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from matplotlib.pylab import mpl
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


class PlotStudioMenu():
    def __init__(self, root):
        self.leftx = 640

        self.menubar = tk.Menu(root)  # 创建菜单栏

        # 创建“文件”下拉菜单
        filemenu = tk.Menu(self.menubar, tearoff=0)
        filemenu.add_command(label="打开", command=lambda: self.file_open(root))
        # filemenu.add_command(label="新建", command=self.file_new)
        # filemenu.add_command(label="保存", command=self.file_save)
        filemenu.add_separator()
        filemenu.add_command(label="退出", command=root.quit)

        # 创建“帮助”下拉菜单
        helpmenu = tk.Menu(self.menubar, tearoff=0)
        helpmenu.add_command(label="关于", command=self.help_about)

        # 将前面三个菜单加到菜单栏
        self.menubar.add_cascade(label="文件", menu=filemenu)
        # self.menubar.add_cascade(label="编辑", menu=editmenu)
        self.menubar.add_cascade(label="帮助", menu=helpmenu)

        # 最后再将菜单栏整个加到窗口 root
        root.config(menu=self.menubar)

    def file_open(self, root):
        file_name = tk.filedialog.askopenfilename()
        wb = xw.Book(file_name)
        sht = wb.sheets[0]
        data_in = sht.range('A1').expand().value
        data = np.array(data_in)

        data_t = data.T
        X = data_t[0][1:]
        Y = []
        for i in range(1, data_t.shape[0]):
            Y.append([float(x) for x in data_t[i][1:]])
        y_max = np.array(Y).max()
        y_min = np.array(Y).min()
        x_max = len(X)

        row = data.shape[0]  # 11
        col = data.shape[1]  # 4

        # 添加treeview控件，预览读入的excel数据表格
        tree = ttk.Treeview(root, show="headings")
        tree["columns"] = tuple(data[0])
        for i in tuple(data[0]):
            tree.column(i, width=80, anchor='c')
            tree.heading(i, text=i)

        # 逐条添加数据
        for i in range(1, row):
            tree.insert('', 'end', values=tuple(data[i]))

        # 添加滚动条，为浏览数据
        vsb = ttk.Scrollbar(tree, orient='vertical', command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(tree, orient='horizontal', command=tree.xview)
        hsb.pack(side='bottom', fill='x')
        tree.configure(xscrollcommand=hsb.set)

        # 放置Treeview控件
        tree.place(x=20, y=20, width=600, height=500)

        title = tk.StringVar()
        title_label = tk.Label(root, text="图片标题：").place(x=self.leftx, y=20)
        title_text = tk.Entry(root, textvariable=title, width=20).place(x=self.leftx + 90, y=20)

        x_label = tk.StringVar()
        xlabel = tk.Label(root, text="横坐标轴标题：").place(x=self.leftx, y=50)
        x_text = tk.Entry(root, textvariable=x_label, width=20).place(x=self.leftx+90, y=50)
        y_label = tk.StringVar()
        ylabel = tk.Label(root, text="纵坐标轴标题：").place(x=self.leftx, y=80)
        y_text = tk.Entry(root, textvariable=y_label, width=20).place(x=self.leftx+90, y=80)

        x_interval = tk.IntVar()
        x_interval_label = tk.Label(root, text="横坐标轴间隔：").place(x=self.leftx, y=130)
        xinterval = tk.Scale(root, from_=1, to=len(X) - 1, orient='horizontal', variable=x_interval, length=140).place(
            x=self.leftx+90,
            y=110)
        # x_interval.set(50)
        y_interval = tk.IntVar()
        y_interval_label = tk.Label(root, text="纵坐标轴间隔：").place(x=self.leftx, y=180)
        yinterval = tk.Scale(root, from_=1, to=100, orient='horizontal', variable=y_interval, length=140).place(x=self.leftx+90,
                                                                                                                y=160)
        x_lim_min = tk.IntVar()
        x_lim_max = tk.IntVar()
        xlim_label = tk.Label(root, text="横坐标范围：").place(x=self.leftx, y=235)
        xlim_min = tk.Entry(root, textvariable=x_lim_min, width=7).place(x=self.leftx+90, y=235)
        x_lim_min.set(1)
        xlim_label2 = tk.Label(root, text="—").place(x=self.leftx+142, y=235)
        xlim_max = tk.Entry(root, textvariable=x_lim_max, width=7).place(x=self.leftx+162, y=235)
        x_lim_max.set(len(X))

        y_lim_min = tk.StringVar()
        y_lim_max = tk.StringVar()
        ylim_label = tk.Label(root, text="纵坐标范围：").place(x=self.leftx, y=265)
        ylim_min = tk.Entry(root, textvariable=y_lim_min, width=7).place(x=self.leftx+90, y=265)
        y_lim_min.set(y_min)
        ylim_label2 = tk.Label(root, text="—").place(x=self.leftx+142, y=265)
        ylim_max = tk.Entry(root, textvariable=y_lim_max, width=7).place(x=self.leftx+162, y=265)
        y_lim_max.set(y_max)

        draw_curve = tk.Button(root, text='绘制图像',
                               command=lambda: drawCurveClick(data_in, title, x_label, y_label, x_interval, y_interval,
                                                              x_lim_min, x_lim_max, y_lim_min, y_lim_max,
                                                              root)).place(x=self.leftx, y=450)

    def help_about(self):
        messagebox.showinfo('关于 PlotStudio',
                            '作者：海江 \n PlotStudio verion 1.0.0 \n 欢迎反馈信息至 \n 微信公众平台 ResearchStudio ')  # 弹出消息提示框
        messagebox._show()


def ColorTransfer(value):
    """
    RGB和16进制颜色之间相互转换
    :param value:
    :return: [r,g,b]或
    """
    digit = list(map(str, range(10))) + list("ABCDEF")
    if isinstance(value, tuple):
        string = '#'
        for i in value:
            a1 = i // 16
            a2 = i % 16
            string += digit[a1] + digit[a2]
        return string
    elif isinstance(value, str):
        a1 = digit.index(value[1]) * 16 + digit.index(value[2])
        a2 = digit.index(value[3]) * 16 + digit.index(value[4])
        a3 = digit.index(value[5]) * 16 + digit.index(value[6])
        return (a1, a2, a3)


def create_matplotlib(data_in, title, x_label, y_label, x_interval, y_interval, x_lim_min, x_lim_max, y_lim_min,
                      y_lim_max):
    plt.cla()
    data = np.array(data_in).T
    X = data[0][1:]

    Y = []
    for i in range(1, data.shape[0]):
        Y.append([float(x) for x in data[i][1:]])

    labels = data_in[0][1:]

    fig = plt.gcf()
    for i in range(len(Y)):
        plt.plot(X, Y[i], label=labels[i], linewidth=0.5)

    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.title(title, fontsize='large')
    plt.ylabel(y_label)
    plt.xlabel(x_label)
    plt.xticks(X[0:len(X): x_interval])
    # plt.yticks(range(0, int(max_y)+1, y_interval))
    # plt.yticks(range(0, len(Y[0]), y_interval))
    plt.xlim(X[x_lim_min - 1], X[x_lim_max - 1])
    plt.ylim(y_lim_min, y_lim_max)

    plt.legend()
    return fig


def drawCurveClick(data_in, title, x_label, y_label, x_interval, y_interval, x_lim_min, x_lim_max, y_lim_min, y_lim_max,
                   root):
    fig = create_matplotlib(data_in, title.get(), x_label.get(), y_label.get(), x_interval.get(), y_interval.get(),
                            x_lim_min.get(), x_lim_max.get(), float(y_lim_min.get()),
                            float(y_lim_max.get()))
    mat_top = tk.Toplevel(root)

    canvas = FigureCanvasTkAgg(fig, mat_top)
    canvas.draw()
    canvas.get_tk_widget().pack(side='top', fill='both', expand=1)

    toolbar = NavigationToolbar2Tk(canvas, mat_top)
    toolbar.update()
    canvas._tkcanvas.pack(side='top', fill='both', expand=1)


def main():
    root = tk.Tk()
    root.title("基于python的折线图生成工具")
    root.geometry("900x550")

    menu = PlotStudioMenu(root)

    root.mainloop()


if __name__ == '__main__':
    main()
