import tkinter as tk
from tkinter import messagebox
import run

window=tk.Tk()
window.title('询价单自动生成脚本v1.1')
window.geometry('400x300')



l1=tk.Label(window,
   font=('宋体',16),
   text='陪标公司1名称：'
   )
l1.pack()

input1=tk.Entry(window,
       font=('宋体',16),
       width=20
       )
input1.pack()

l2=tk.Label(window,
   font=('宋体',16),
   text='陪标公司2名称：'
   )
l2.pack()

input2=tk.Entry(window,
       font=('宋体',16),
       width=20
       )
input2.pack()

def fun():
   company1=input1.get()
   company2=input2.get()
   rows=run.getExcelRow()
   flag=1
   while flag<=rows:
      info=run.readExcelData(flag)
      run.editDocxWin(info)
      run.editDocxCompany1(info,company1)
      run.editDocxCompany2(info,company2)
      flag+=1

def callback():
    if messagebox.askyesno('Verify', '确定陪标公司名称无误？'):
        messagebox.showwarning('确认', '询价单生成完毕！！！')
        fun()
    else:
        messagebox.showinfo('取消', '请重新输入陪标公司名称')

btn=tk.Button(window,
    font=('宋体',16),
    text='生成',
    height=2,
    width=10,
    activeforeground='grey',
    command=callback
    )
btn.pack()



window.mainloop()