import pyodbc
import glob
import os
from tkinter import *
import sys
from tkinter import messagebox
import win32com.client

def update():
    rno_searched=str(e3.get())
    course_selected=str(menu2.get())
    tab1=course_selected
    filendat=e41.get()
    filnmdat=e42.get()
    filgndat=e43.get()
    filmndat=e44.get()
    print("Update "+tab1+" set EN='"+filendat+"' where RN="+rno_searched)
    inputdb.execute("Update "+tab1+" set EN='"+filendat+"' where RN="+rno_searched)
    inputdb.execute("Update "+tab1+" set NM='"+filnmdat+"' where RN="+rno_searched)
    inputdb.execute("Update "+tab1+" set GN='"+filgndat+"' where RN="+rno_searched)
    inputdb.execute("Update "+tab1+" set MN='"+filmndat+"' where RN="+rno_searched)
    inputdb.commit()
    messagebox.showinfo("Updated", "Performed Changes Were Successful !!!")

def backtoroot2():
    try:
        access.DoCmd.CloseDatabase
        access.Quit()
    except:
        pass
    root3.destroy()
    root.deiconify()

def but4():
    inputdb.close()
    try:
        access.DoCmd.CloseDatabase
        access.Quit()
    except:
        pass
    sys.exit(0)

def reportopenprint():
    global access
    access = win32com.client.Dispatch("Access.Application")
    database = access.OpenCurrentDatabase(filename)
    access.DoCmd.OpenReport("rep",0)

def but3print():
    global course_selected,access
    try:
        access.DoCmd.CloseDatabase
        access.Quit()
    except:
        pass
    rno_searched=str(e3.get())
    course_selected=str(menu2.get())
    tab1=course_selected
    query="SELECT * INTO TRES FROM "+tab1+" as tab1 where tab1.RN="+rno_searched+" ORDER BY tab1.RN;"
    print (query)
    try:
        tabcur1.execute("Drop table TRES;")
        inputdb.commit()
    except:
        pass
    tabcur1.execute(query)
    inputdb.commit()
    tabcur1.execute("select * from TRES")
    rres=tabcur1.fetchall()
    if rres==[]:
        messagebox.showerror("Error 404","No Matching Records Found !!")
    else:
        reportopenprint()

def reportopen():
    global access
    access = win32com.client.Dispatch("Access.Application")
    database = access.OpenCurrentDatabase(filename)
    access.DoCmd.OpenReport("rep",2)

def but3():
    global course_selected,access
    try:
        access.DoCmd.CloseDatabase
        access.Quit()
    except:
        pass
    rno_searched=str(e3.get())
    course_selected=str(menu2.get())
    tab1=course_selected
    query="SELECT * INTO TRES FROM "+tab1+" as tab1 where tab1.RN="+rno_searched+" ORDER BY tab1.RN;"
    print (query)
    try:
        tabcur1.execute("Drop table TRES;")
        inputdb.commit()
    except:
        pass
    tabcur1.execute(query)
    inputdb.commit()
    tabcur1.execute("select * from TRES")
    rres=tabcur1.fetchall()
    if rres==[]:
        messagebox.showerror("Error 404","No Matching Records Found !!")
    else:
        reportopen()

def but5main():
    global course_selected
    try:
        access.DoCmd.CloseDatabase
        access.Quit()
    except:
        pass
    rno_searched=str(e3.get())
    course_selected=str(menu2.get())
    tab1=course_selected
    query="SELECT * INTO TRES FROM "+tab1+" as tab1 where tab1.RN="+rno_searched+" ORDER BY tab1.RN;"
    print (query)
    try:
        tabcur1.execute("Drop table TRES;")
        inputdb.commit()
    except:
        pass
    tabcur1.execute(query)
    inputdb.commit()
    tabcur1.execute("select * from TRES")
    rres=tabcur1.fetchall()
    if rres==[]:
        messagebox.showerror("Error 404","No Matching Records Found !!")
    else:
        but5()

def but5():
    global e41,e42,e43,e44
    root4=Toplevel(root)
    root4.title("Update Details")
    l41=Label(root4,text="Enrollment Number").grid(row=1,column=1,sticky="news",padx=10,pady=10)
    l42=Label(root4,text="Name").grid(row=2,column=1,sticky="news",padx=10,pady=10)
    l43=Label(root4,text="Father's Name").grid(row=3,column=1,sticky="news",padx=10,pady=10)
    l44=Label(root4,text="Mother's Name").grid(row=4,column=1,sticky="news",padx=10,pady=10)
    e41=Entry(root4)
    e42=Entry(root4)
    e43=Entry(root4)
    e44=Entry(root4)
    e41.grid(row=1,column=2,sticky="news",padx=10,pady=10)
    e42.grid(row=2,column=2,sticky="news",padx=10,pady=10)
    e43.grid(row=3,column=2,sticky="news",padx=10,pady=10)
    e44.grid(row=4,column=2,sticky="news",padx=10,pady=10)
    tabcur1.execute("Select EN,NM,GN,MN from TRES")
    fildat=tabcur1.fetchall()
    e41.delete(0,"end")
    e41.insert(0,fildat[0][0])
    e42.delete(0,"end")
    e42.insert(0,fildat[0][1])
    e43.delete(0,"end")
    e43.insert(0,fildat[0][2])
    e44.delete(0,"end")
    e44.insert(0,fildat[0][3])
    b41=Button(root4,text="Update",command=update)
    b41.grid(row=5,column=1,sticky="news",padx=10,pady=10)
    b42=Button(root4,text="Cancel",command=root4.destroy)
    b42.grid(row=5,column=2,sticky="news",padx=10,pady=10)

def but2():
    global e3,root3
    root3=Toplevel(root)
    root.withdraw()
    root3.title("Roll Number")
    l3=Label(root3,text="Enter Roll Number").grid(row=1,column=1,sticky="news",padx=10,pady=10)
    e3=Entry(root3)
    e3.grid(row=1,column=2,sticky="news",padx=10,pady=10)
    b3=Button(root3,text="Change",command=but5main)
    b3.grid(row=2,column=1,sticky="news",padx=10)
    b3=Button(root3,text="Print Preview",command=but3)
    b3.grid(row=2,column=2,sticky="news",padx=10)
    b33=Button(root3,text="Back",command=backtoroot2)
    b33.grid(row=3,column=1,sticky="news",padx=10,pady=10)
    b32=Button(root3,text="Exit",command=but4)
    b34=Button(root3,text="Print",command=but3print)
    b34.grid(row=3,column=2,rowspan=2,sticky="news",padx=10,pady=10)
    b32.grid(row=4,column=1,sticky="news",padx=10,pady=10)
    root3.mainloop()

def but1(*args):
    global menu1,menu2,root2,tabcur1,inputdb,filename,options2
    tablenames1=[]
    filename=str(os.getcwd())+'\\'+str(menu1.get())
    try:
        inputdb = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+filename+';')
    except:
        messagebox.showerror("Fatal Error !!","Couldn't Open Database. Exiting !")
        sys.exit(0)
    tabcur1=inputdb.cursor()
    temptablenames1=tabcur1.tables()
    for i in temptablenames1:
        if i[3]=='TABLE':
            tablenames1.append(i[2])
    options2 = [x for x in tablenames1 if x != "TRES"]
    drop2 = OptionMenu(root,menu2,*options2).grid(row=2,column=2,sticky="news",padx=10,pady=10)
    menu2.set(options2[0])

root=Tk()
root.title("Choose Database & Tables")
l1=Label(text="Select Database").grid(row=1,column=1,sticky="news",padx=10,pady=10)
options1 = [x[2:] for x in glob.glob('./*.accdb')]+[x[2:] for x in glob.glob('./*.mdb')]
if options1==[]:
    messagebox.showerror("Fatal Error !!","Couldn't Find Any Databases. Exiting !")
    sys.exit(0)
menu1= StringVar()
menu1.set(options1[0])
drop1 = OptionMenu(root,menu1,*options1 ).grid(row=1,column=2,sticky="news",padx=10,pady=10)
menu1.trace("w",but1)
l2=Label(root,text="Select Table").grid(row=2,column=1,sticky="news",padx=10,pady=10)
options2 = ["---"]
menu2= StringVar()
menu2.set(options2[0])
drop2 = OptionMenu(root,menu2,*options2).grid(row=2,column=2,sticky="news",padx=10,pady=10)
b2=Button(root,text="Continue",command=but2)
b2.grid(row=3,column=1,columnspan=2,sticky="news",padx=10,pady=10)
root.mainloop()