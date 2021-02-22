import tkinter as tk
import tkinter.messagebox
import cacu
import Excel

class excpt():
    def __init__(self):
        self.root = tkinter.Tk()
        self.root.withdraw()
        self.cacu = cacu.cacu()

    def try_start(self,start_hours, start_minuit):
        try :
            int(start_hours)
            int(start_minuit)
        except:
            tkinter.messagebox.showinfo('缺少参数','请将参数输入完整')
        else:
            self.cacu.get_start_time(int(start_hours),int(start_minuit))



    def try_stop(self, stop_hours, stop_minuit):
        try :
            int(stop_hours)
            int(stop_minuit)
        except:
            tkinter.messagebox.showinfo('缺少参数','请将参数输入完整')
        else:
            self.cacu.get_stop_time(int(stop_hours),int(stop_minuit))

    def try_price(self,price,isTriple):
        try:
            int(price)
        except:
            tkinter.messagebox.showinfo('缺少参数','请将参数输入完整')
        else:
            self.cacu.cacuresult(int(price),isTriple)

    def try_write(self):
        try:
            self.cacu.writer_to_excel()
        except:
            tkinter.messagebox.showinfo('缺少参数', '请将参数输入完整')

    def try_relax(self,relax):
        try:
            self.cacu.get_relax_time(int(relax))
        except:
            self.cacu.get_relax_time(0)

class excel():
    def __init__(self):
        self.excle_data = Excel.Execl()
        self.cacu = cacu.cacu()

    def passed(self):
        self.excle_data.find_data()
        j = 0
        for i in range(0,self.excle_data.lenth,2):
            self.cacu.writer_to_excel(self.excle_data.line_v[i],self.excle_data.line_v[i+1],self.excle_data.relaxtime,self.excle_data.locale_price[j])
            j = j + 1




class UI():

    def __init__(self):
        self.Tk = tk.Tk()
        self.Tk.geometry("450x200")
        self.Tk.title('小时工工资计算器')

        self.excpt = excpt()
        self.cacu = cacu.cacu()
        self.excel = excel()

        self.showstart()
        self.showstop()
        self.insert_Excel()
        self.caculator()
        self.relax_time()
        self.read_Excel()


        self.Tk.mainloop()



    def showstart(self):
        self.starth_t = tk.StringVar()
        self.startm_t = tk.StringVar()

        start_l = tk.Label(self.Tk, text='开始时间：')
        starth = tk.Entry(self.Tk, width=10,textvariable = self.starth_t)
        start_h = tk.Label(self.Tk, text='时')
        startm = tk.Entry(self.Tk, width=10,textvariable = self.startm_t)
        start_m = tk.Label(self.Tk, text='分')
        OK = tk.Button(self.Tk, width=5, height=1, text='输入', command=lambda:self.excpt.try_start(starth.get(),startm.get()))

        start_l.pack()
        start_l.place(x=5, y=5, anchor='nw')
        starth.pack()
        starth.place(x=70, y=5, anchor='nw')
        start_h.pack()
        start_h.place(x=150, y=5, anchor='nw')
        startm.pack()
        startm.place(x=180, y=5, anchor='nw')
        start_m.pack()
        start_m.place(x=250, y=5, anchor='nw')

        OK.pack()
        OK.place(x=400, y=5, anchor='nw')

    def showstop(self):
        self.stoph_t = tk.StringVar()
        self.stopm_t = tk.StringVar()

        stop = tk.Label(self.Tk, text='结束时间：')
        stoph = tk.Entry(self.Tk, width=10,textvariable = self.stoph_t)
        stop_h = tk.Label(self.Tk, text='时')
        stopm = tk.Entry(self.Tk, width=10,textvariable = self.stopm_t)
        stop_m = tk.Label(self.Tk, text='分')
        stopOK = tk.Button(self.Tk, width=5, height=1, text='输入', command=lambda:self.excpt.try_stop(stoph.get(),stopm.get()))

        stop.pack()
        stop.place(x=5, y=40, anchor='nw')
        stoph.pack()
        stoph.place(x=70, y=40, anchor='nw')
        stop_h.pack()
        stop_h.place(x=150, y=40, anchor='nw')
        stopm.pack()
        stopm.place(x=180, y=40, anchor='nw')
        stop_m.pack()
        stop_m.place(x=250, y=40, anchor='nw')

        stopOK.pack()
        stopOK.place(x=400, y=40, anchor='nw')

    def insert_Excel(self):
        insert = tk.Button(self.Tk,text='输入Excel', width=50, height=1,command = self.excpt.try_write)

        insert.pack()
        insert.place(x=225, y=135, anchor='n')

    def relax_time(self):
        relax_t = tk.StringVar()

        insert_l = tk.Label(self.Tk,text = '休息时间')
        insert = tk.Entry(self.Tk,textvariable = relax_t)

        insert_b = tk.Button(self.Tk,text = '输入',width = 5,height = 1,command = lambda:self.excpt.try_relax(insert.get()))

        insert_b.pack()
        insert_b.place(x=400,y=72,anchor = 'nw')
        insert_l.pack()
        insert_l.place(x=5,y=72,anchor = 'nw')
        insert.pack()
        insert.place(x=70,y=72,anchor = 'nw')

    def caculator(self):
        self.Triple = 1
        self.price_t = tk.StringVar()

        price = tk.Label(self.Tk,text = '小时工资')
        price_e = tk.Entry(self.Tk,textvariable = self.price_t)
        holiday = tk.Checkbutton(self.Tk,text = '当日三倍工资',variable = self.Triple)
        result = tk.Button(self.Tk,text='计算工资', width=8, height=1, command=lambda: self.excpt.try_price(price_e.get(),self.Triple))

        result.pack()
        result.place(x=380,y=105,anchor = 'nw')
        price.pack()
        price.place(x=5,y=105,anchor = 'nw')
        holiday.pack()
        holiday.place(x=210,y=105,anchor='nw')
        price_e.pack()
        price_e.place(x=70,y=105,anchor = 'nw')

    def read_Excel(self):
        insert = tk.Button(self.Tk,text='读取Excel表', width=50, height=1, command = self.excel.passed)

        insert.pack()
        insert.place(x=225, y=165, anchor='n')

