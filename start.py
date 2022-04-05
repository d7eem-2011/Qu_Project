#!/usr/bin/python
# -*- coding=utf-8 -*-
import unittest
import sys
import pyarabic.number as nb
import os
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox as mb
from module3 import *
from num2words import num2words



# # Language Support.
# print(num2words(36, lang ='es'))


qux = Tk()

photo = PhotoImage(file=r"image\qu1.png")
labelImg = Label(qux, width=377, height=140, image=photo, background="lightblue", compound=TOP).pack(side=TOP)

qux.title("-الصفحة الرئيسية-")
qux.geometry("600x390")
qux.resizable(False, False)
qux.config(background="lightblue")
font = ('Helvetica')
style = ttk.Style()
ttk.Style().configure("TButton",  relief="flat",background="#03A6A6", font=font)
style.map("C.TButton",
    foreground=[('pressed', 'red'), ('active', 'blue')],
    background=[('pressed', '!disabled', 'black'), ('active', 'white')])
style.configure("TOptionMenu", background=[('pressed', '!disabled', 'black'), ('active', 'white')])



def close():
    qux.destroy()


def oneWindow():
    boneWindow = Toplevel(qux)

    def close():
        boneWindow.destroy()

    def open1():
        fileopen1 = filedialog.askdirectory(initialdir="/", title="اختر ملف الحفظ: ")
        locaPath.insert(0, fileopen1)


    def call():
        res = mb.askquestion('اغلاق النافذة', 'الخروج دون الحفظ')
        if res == 'yes':
            boneWindow.destroy()
        else:
            mb.showinfo('العوده', 'البقاء في النافذة')

    def open_tashkel():
        global fileopen_tashkel
        fileopen_tashkel = filedialog.askopenfilename(initialdir="/", title="اختر ملف التشكيل:", filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))
        tashkelpath.insert(0, fileopen_tashkel)
        # labelFile = Label(boneWindow, text=fileopen_tashkel,
        #                   background="white")
        # labelFile.pack(side=TOP)

    def open_ketat():
        global fileopen_ketat
        fileopen_ketat = filedialog.askopenfilename(
            initialdir="/", title="اختر ملف الخطط:", filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))

        katatpath.insert(0, fileopen_ketat)
        # labelFile = Label(boneWindow, text=fileopen_ketat, background="white")
        # labelFile.pack(side=TOP)

    def task_run():
        docfil = r'World Docement\template v4.docx'
        doctash = r"World Docement\تشكيل - مذكرات عرض للمجلس.docx"
        docad = r"World Docement\اعلان الجلسة.docx"
        docnote = r"World Docement\اعلان الخطط.docx"
        docsum = r"World Docement\ملخص الخطط والتشكيل لوحدة الامن الفكري.docx"



        if locaPath.get():
            try:

                if tashkelpath.get() and katatpath.get():

                    tashkel = tashkelpath.get()
                    katat = katatpath.get()
                    name_clicked = clicked.get()
                    name_clicked = 'الجلسة '+name_clicked
                    datein = str(cal3.get())
                    try:
                        savePath = locaPath.get()
                    except NameError as n:
                        mb.showinfo('لم يتم اختيار مسار الحفظ', ' الرجاء اختيرا مسار الحفظ \n')



                    task_do(katat, tashkel, docfil, doctash, name_clicked, datein, realPath=savePath)
                    task_do2(katat, tashkel, docad, name_clicked, datein, realPath=savePath)
                    task_do3(katat, tashkel, docnote, name_clicked, datein, realPath=savePath)
                    task_do4(katat, tashkel, docsum, name_clicked, datein, realPath=savePath)
                    # time.sleep(3)

                    mb.showinfo("تم الانتهاء من الحفظ",
                                "تم الانتهاء من انشاء المجلدات,\n تحقق من مسار الحفظ.")
                else:
                    mb.showinfo('لم يتم اختيار الملفات', ' الرجاء اختيرا التشكيل و الخطط \n')
            except Exception as a:
                mb.showerror(" Warning ", f" {a} ")
        else:
            mb.showinfo('لم يتم اختيار مسار الحفظ', ' الرجاء اختيرا مسار الحفظ \n')


    def checkSave():

        if  locaPath.get():
            path = os.path.realpath(locaPath.get())
            os.startfile(path)
        else:
            mb.showinfo('لم يتم اختيار مسار الحفظ', ' الرجاء اختيرا مسار الحفظ \n')

    boneWindow.title("تسجيل جلسة جديدة")
    boneWindow.geometry("700x500")
    boneWindow.resizable(False, False)
    boneWindow.config(background="lightblue")

    labelImg = Label(boneWindow, width=377, height=140, image=photo, background="lightblue", compound=TOP)
    labelImg.pack(side=TOP)

    l1 = Label(boneWindow, text="حدد رقم الجلسة", width=15, height=1, background="lightblue", compound=TOP)
    l1.place(x=570, y=180)

    l2 = Label(boneWindow, text="ادخل تاريخ الجلسة", width=15, height=1, background="lightblue", compound=TOP)
    l2.place(x=350, y=180)

    # options = [*range(1,31)]
    #
    # clicked = StringVar()
    # clicked.set(options[0])
    # drop = OptionMenu(boneWindow, clicked, *options)
    # drop.place(x=295, y=175)

    options = []
    for i in range(1, 31):
        
        if i == 1:
            num = u'الاولى'
        else:
            # num = num2words(i, lang='ar', ordinal=True, to='ordinal_num')
            num = nb.number2ordinal(i, feminin=True)
        options.append(num)

    clicked = StringVar()
    clicked.set(options[0])
    drop = ttk.OptionMenu(boneWindow, clicked, *options)
    drop.place(x=470, y=180)

    # cal = DateEntry(boneWindow, width=15, background='black', foreground='white', borderwidth=2)
    # cal.place(x=240, y=220)

    cal3 = ttk.Entry(boneWindow, width=20, justify=RIGHT)
    cal3.place(x=230, y=180)

    tashkelButton = ttk.Button(boneWindow, text='اختيار ملف التشكيل', command=open_tashkel, style="C.TButton").place(x=570, y=220)

    katatButton = ttk.Button(boneWindow, text="اختيار ملف الخطط", command=open_ketat, style="C.TButton").place(x=570, y=260)

    tashkelpath = ttk.Entry(boneWindow, width=60, justify=RIGHT)
    tashkelpath.place(x=200, y=220)

    katatpath = ttk.Entry(boneWindow, width=60, justify=RIGHT)
    katatpath.place(x=200, y=260)

    l3 = ttk.Label(boneWindow, text=":ملاحظة قبل الحفظ* \n حدد رقم الجلسة - \n ادخل تاريخ الجلسة - \n اختيار ملف التشكيل - \n اختيار ملف الخطط -",width=20, justify=RIGHT , borderwidth=20, background="#FFFFFF")
    l3.place(x=15, y=200)

    saveButton = ttk.Button(boneWindow,width=20 , text="حفظ", command=task_run, style='C.TButton').place(x=360, y=310)

    b4 = ttk.Button(boneWindow, text="التحقق من الحفظ", command=checkSave, style='C.TButton').place(x=200, y=310)

    chosePath = ttk.Button(boneWindow, text=" : اختيار ملف الحفظ ", command=open1)
    chosePath.place(x=570, y=370)
    locaPath = ttk.Entry(boneWindow, width=60, justify=RIGHT, valu=None)
    locaPath.place(x=200, y=370)

    buttonClose = ttk.Button(boneWindow, text="اغلاق", command=call, style='C.TButton').place(x=355, y=420)


b1 = ttk.Button(qux, text="تسجيل جلسة جديدة", command=oneWindow, style="C.TButton", width=40).place(x=120, y=215)


def fiveWindow():
    bfiveWindow = Toplevel(qux)



    bfiveWindow.title("اعداد البرنامج")
    bfiveWindow.geometry("600x480")
    bfiveWindow.resizable(False, False)
    bfiveWindow.config(background="lightblue")

    labelImg = Label(bfiveWindow, width=377, height=140, image=photo, background="lightblue", compound=TOP).pack(side=TOP)

    b1Label = Label(bfiveWindow, text="""
    مصمم البرنامج عبدالرحمن سالم الشحيتان 
    للتواصل:
    Email:abd2980@outlook.com
    phone:+966 533111535
    """,
    justify=RIGHT, background="lightblue", font=font).place(x=170,y=220)





    b4 = Button(bfiveWindow, text="اغلاق", width=10,
                height=1, command=close).place(x=255, y=420)


button5 = ttk.Button(qux, text="عن البرنامج", command=fiveWindow, style="C.TButton").place(x=80, y=320)


#buttonClose = Button(qux, text="اغلاق",borderwidth=5, command=close, relief=FLAT, bg="#03A6A6", font=font).place(x=255, y=320)
buttonClose = ttk.Button(qux, text="اغلاق", command=close, style="C.TButton", width=25).place(x=200, y=320)

if __name__ == '__main__':
    qux.mainloop()
