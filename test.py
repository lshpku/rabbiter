import tkinter
from tkinter import ttk


win = tkinter.Tk()
win.title("Kahn Software v1")    # #窗口标题
win.geometry("500x300+200+20")   # #窗口位置500后面是字母x

'''
下拉菜单
'''
xVariable = tkinter.StringVar()     # #创建变量，便于取值


class Combobox():
    def __init__(self, master=None):
        self.values = tkinter.StringVar()
        self.cbox = ttk.Combobox(master, textvariable=self.values)

    def pack(self, cnf={}, **kw):
        self.cbox.pack(cnf, **kw)
        return self

    def set(self, values: list):
        self.cbox['value'] = values
        self.cbox.current(0)

    def get(self) -> str:
        return self.cbox.get()


com = ttk.Combobox(win, textvariable=xVariable)     # #创建下拉菜单
com.pack()     # #将下拉菜单绑定到窗体
com["value"] = ("河北", "河南", "山东")    # #给下拉菜单设定值
com.current(2)    # #设定下拉菜单的默认值为第3个，即山东


def xFunc(event):
    print(com.get())            # #获取选中的值方法1
    print(xVariable.get())      # #获取选中的值方法2


com.bind("<<ComboboxSelected>>", xFunc)     # #给下拉菜单绑定事件

cbox = Combobox(win)
cbox.set(['fsd', 'dfsfs', 'dfsdf'])

win.mainloop()
