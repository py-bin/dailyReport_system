# -*- coding: utf-8 -*-
"""
Created on Tue Apr 24 11:54:59 2018

@author: bin
"""

import oscpc
import kuandai
import iptv
import tisu
import tkinter as tk
import tkinter.messagebox
from PIL import Image, ImageTk
import tkinter.font as tkFont

root = tk.Tk()
root.title("日报数据处理系统")
root.geometry('800x600')
root.resizable(False, False)

txt = tk.Text(root, height=10, width=70)
txt.place(x=50,y=300)

def osc_read():
    oscpc.read_xt()
    txt.insert('end','已读入O2O数据！\n')
    txt.focus_force()
    txt.see(tk.END)
def osc_pdrq():
    rqtxt = oscpc.pdrq()
    txt.insert('end',rqtxt+'\n')
    txt.focus_force()
    txt.see(tk.END)
def osc_process():
    oscpc.process()
    txt.insert('end','O2O数据已更新\n')    
    txt.focus_force()
    txt.see(tk.END)
def osc_total():
    osc_read()
    osc_pdrq()
    osc_process()

def kd_read():
    kuandai.read_kd()
    txt.insert('end','已读入宽带数据！\n')
    txt.focus_force()
    txt.see(tk.END)
def kd_pdqs():
    txt.insert('end',kuandai.pdqs())
    txt.focus_force()
    txt.see(tk.END)
def kd_process():
    kuandai.process()
    txt.insert('end','宽带数据已更新\n')    
    txt.focus_force()
    txt.see(tk.END)
def kd_total():
    kd_read()
    kd_pdqs()
    kd_process()

def itv_read():
    iptv.read_itv()
    txt.insert('end','已读入IPTV数据！\n')
    txt.focus_force()
    txt.see(tk.END)
def itv_pdqs():
    txt.insert('end',iptv.pdqs())
    txt.focus_force()
    txt.see(tk.END)
def itv_process():
    iptv.process()
    txt.insert('end','IPTV数据已更新\n')    
    txt.focus_force()
    txt.see(tk.END)
def itv_total():
    itv_read()
    itv_pdqs()
    itv_process()
    
def ts_read():
    tisu.read_ts()
    txt.insert('end','已读入提速数据！\n')
    txt.focus_force()
    txt.see(tk.END)    
def ts_pdqs():
    txt.insert('end',tisu.pdqs())
    txt.focus_force()
    txt.see(tk.END)
def ts_process():
    tisu.process()
    txt.insert('end','提速数据已更新\n')    
    txt.focus_force()
    txt.see(tk.END)    
def ts_total():
    ts_read()
    ts_pdqs()
    ts_process()    

#目录
menubar = tk.Menu(root)
#协同目录
o2omenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label=' 协同O2O ', menu=o2omenu)
o2omenu.add_command(label='打开数据', command=osc_read)
o2omenu.add_command(label='检查日期', command=osc_pdrq)
o2omenu.add_command(label='处理数据', command=osc_process)
o2omenu.add_separator()
o2omenu.add_command(label='协同O2O一键处理', command=osc_total)
#宽带目录
kdmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label=' 宽带数据 ', menu=kdmenu)
kdmenu.add_command(label='打开数据', command=kd_read)
kdmenu.add_command(label='检查缺数', command=kd_pdqs)
kdmenu.add_command(label='处理数据', command=kd_process)
kdmenu.add_separator()
kdmenu.add_command(label='宽带一键处理', command=kd_total)
#iptv目录
iptvmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label=' IPTV数据 ', menu=iptvmenu)
iptvmenu.add_command(label='打开数据', command=itv_read)
iptvmenu.add_command(label='检查缺数', command=itv_pdqs)
iptvmenu.add_command(label='处理数据', command=itv_process)
iptvmenu.add_separator()
iptvmenu.add_command(label='IPTV一键处理', command=itv_total)
#提速目录
tsmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label=' 提速数据 ', menu=tsmenu)
tsmenu.add_command(label='打开数据', command=ts_read)
tsmenu.add_command(label='检查缺数', command=ts_pdqs)
tsmenu.add_command(label='处理数据', command=ts_process)
tsmenu.add_separator()
tsmenu.add_command(label='提速一键处理', command=ts_total)

root.config(menu=menubar)

#监控区域
ziti1 = tkFont.Font(size=15, weight=tkFont.BOLD)
lb_jk = tk.Label(root,text='数据监控',font=ziti1)
lb_jk.place(x=320,y=40)

root.mainloop()

#==============================================================================
# © 2018 Zheng, All Rights Reserved
#==============================================================================
