# load libraries
import os
from datetime import date
import pyarabic.number
from num2words import num2words
from docxtpl import DocxTemplate
import pandas as pd


#
# #for the time
# v=time.time()
# t=time.ctime(v)


# docfil = r'C:\Users\D7eem\PycharmProjects\Qu_Project\World Docement\template v4.docx'
#
# xlxfil = r'C:\Users\D7eem\Desktop\OneDrive_3_10-5-2020\خطط الجلسة الثالثة.xlsx'
#
# doc = DocxTemplate(docfil)


def ketat_def(X):
    ex = pd.read_excel(X, header=0)

    ex = ex.fillna(0)

    stdname = pd.Series(ex['اسم الطالب'])
    ex['الرقم الجامعي'] = ex['الرقم الجامعي'].astype(int)
    stdnumber = pd.Series(ex['الرقم الجامعي'])
    gen = pd.Series(ex['الجنس'])
    cole = pd.Series(ex['الكلية'])
    department = pd.Series(ex['القسم'])
    edu = pd.Series(ex['المرحلة (ماجستير - دكتوراه)'])
    program = pd.Series(ex['البرنامج (التخصص)'])
    numdepart = pd.Series(ex['رقم جلسة القسم'])
    departdate = pd.Series(ex['تاريخ جلسة القسم'])
    numcole = pd.Series(ex['رقم جلسة الكلية'])
    coledate = pd.Series(ex['تاريخ جلسة الكلية'])
    title = pd.Series(ex['عنوان الرسالة'])

    supervisorname = pd.Series(ex['اسم المشرف'])
    supervisoredu = pd.Series(ex['الرتبة العلمية'])

    supervisorname2 = pd.Series(ex['اسم مساعد المشرف'])
    supervisoredu2 = pd.Series(ex['الرتبة العلمية لمساعد المشرف'])

    supervisorgen = pd.Series(ex['جنس المشرف'])

    ordernumber = pd.Series(ex['رقم الطلب'].astype(int))





    total = stdnumber.count()

    context = {'خطط': []}

    switch_std = {
        'ذكر': 'طالب',
        'أنثى': 'طالبة'
    }

    switch_doc = {
        'ذكر': 'دكتور',
        'أنثى': 'دكتورة'
    }

    switch_to = {
        'ذكر': 'ذي',
        'أنثى': 'ذات'
    }

    for i in range(total):
        context['خطط'].append({

            'رقم': pyarabic.number.number2ordinal(i + 1),
            'اسمالطالب': stdname[i],
            'رقمالطالب': stdnumber[i],
            'الكلية': cole[i],
            'تاريخالقسم': departdate[i],
            'تاريخالكلية': coledate[i],
            'رقمجلسكلية': pyarabic.number.number2ordinal(int(numcole[i])),
            'البرنامج': program[i],
            'العنوان': title[i],
            'القسم': department[i],
            'رقمقسم': pyarabic.number.number2ordinal(int(numdepart[i])),
            'التعليم': edu[i],
            'الطالب': switch_std.get(gen[i]),
            'دكتور': switch_doc.get(supervisorgen[i]),
            'اسمالمشرف': supervisorname[i],
            'تعليمالمشرف': supervisoredu[i],
            'اسمالمشرف2': supervisorname2[i],
            'تعليمالمشرف2': supervisoredu2[i],
            'ذي': switch_to.get(gen[i]),
            'رقمالطلب': ordernumber[i]





        })
    return context
    # print(context)
    #
    # doc.render(context)
    # today = str(date.today())
    # OUTPUT = 'output تجربة اعداد محضر'
    # if not os.path.exists(OUTPUT):
    #     os.makedirs(OUTPUT)
    #
    # save_name = today + ' template output v3 .docx'
    # doc.save(OUTPUT + '/' + save_name)
    #
