#coding:utf8
import xlrd
from xlutils.copy import copy
from tkinter import *
import datetime

def destroy_root(event):
    root.destroy()
def Run_time(runtime):
    L4['text']=runtime
def read_excel(x):
    month=Entry_month.get()
    list1sum=[]
    workbook=xlrd.open_workbook('D:111sss.xlsx')
    sheet6_name=workbook.sheet_names()[7]
    sheet6=workbook.sheet_by_name(sheet6_name)
    summ=0
    molist=list(month)
    if(len(molist)==3):
        mostr=molist[0]+molist[1]
        monum=int(mostr)+1
    else:
        monum=int(molist[0])+1
    listx=[]
    for j in range(1,monum):
        listx.append(sheet6.cell(x,j).value)
        summ+=round(sheet6.cell(x,j).value)
    return summ
def alllist(event):
    month=Entry_month.get()
    stime=datetime.datetime.now()
    name='D:'+Entry_name.get()+'.xls'
    list1=list(range(2,29))
    list2=list(range(36,63))
    list3=list(range(70,97))
    list4=list(range(104,132))
    list5=list(range(139,167))
    list6=list(range(174,202))
    list7=list(range(209,237))
    list8=list(range(244,272))
    list9=list(range(279,306))
    list10=list(range(312,335))
    list11=list(range(342,370))
    list12=list(range(377,405))
    listt=[list1,list2,list3,list4,list5,list6,list7,list8,list9,list10,list11,list12]
    list1sum=[]
    list2sum=[]
    list3sum=[]
    list4sum=[]
    list5sum=[]
    list6sum=[]
    list7sum=[]
    list8sum=[]
    list9sum=[]
    list10sum=[]
    list11sum=[]
    list12sum=[]
    listsum=[list1sum,list2sum,list3sum,list4sum,list5sum,list6sum,list7sum,list8sum,list9sum,list10sum,list11sum,list12sum]
    for i in range(12):
        for j in listt[i]:
            listsum[i].append(read_excel(j))
        print(listsum[i])
    #写文件
    workbook=xlrd.open_workbook('D:111sss.xlsx')
    sheet6_name=workbook.sheet_names()[7]
    wb=copy(workbook)
    ws=wb.get_sheet(7)
    ws.write(1,16,month)
    for i in range(12):
        for j in range(len(listt[i])):
            ws.write(j+listt[i][0],16,listsum[i][j])
    wb.save(name)
    etime=datetime.datetime.now()
    runtime=etime-stime
    
    Run_time('一共用了'+str(runtime.seconds)+'秒')

#gui入口
root = Tk()
root.title('excel transformer')
root.geometry('400x200')

L1=Label(root,text='请在框中输入以下信息')
L1.pack()
fram1=Frame(root)
L2=Label(fram1,text='月份:')
Entry_month=Entry(fram1,width=20)
L2.pack(side='left')
Entry_month.pack()
fram1.pack()
fram2=Frame(root)
L3=Label(fram2,text='文件名:')
Entry_name=Entry(fram2,width=20)
Entry_name.bind('<Return>',alllist)
L3.pack(side='left')
Entry_name.pack()
fram2.pack()

fram3=Frame(root)
butt1=Button(fram3,text='确定')
butt2=Button(fram3,text='关闭')
butt1.bind('<Button-1>',alllist)
butt2.bind('<Button-1>',destroy_root)
butt1.pack(side='left')
butt2.pack()
fram3.pack()
L4=Label(fram2)
L4.pack()
root.mainloop()
