import tkinter.messagebox
import UI
import tkinter


def main():
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showinfo(
    '使用说明',
    '    1.输入开始和结束的时间\n\
    2.点击输入\n\
    3.输入小时工资并点击计算\n\
    4.点击写入Excel\n\
    5.读取Excel表功能根据提示一步步操作即可')
    MainUI = UI.UI()


if __name__ == '__main__':
    main()