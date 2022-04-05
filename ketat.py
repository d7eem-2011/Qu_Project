# load libraries
import os
from datetime import date
import pyarabic.number as nb
from num2words import num2words
from docxtpl import DocxTemplate
import pandas as pd
# from module3 import corctnum


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
    context = {'خطط': []}
    try:
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

        Duration_of_study = pd.Series(ex['مدة الدراسة'])







        total = stdnumber.count()



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

        def corctnum(x, y=True):
            if x + 1 == 1:
                return 'الاول'
            else:
                return nb.number2ordinal(x + 1, feminin=y)

        for i in range(total):
            if stdname[i] == 0 or stdnumber[i] == 0:
                continue

            context['خطط'].append({
                'رقم': corctnum(i, y=False),
                'اسمالطالب': stdname[i],
                'رقمالطالب': stdnumber[i],
                'الكلية': cole[i],
                'تاريخالقسم': departdate[i],
                'تاريخالكلية': coledate[i],
                'رقمجلسكلية': corctnum(int(numcole[i])),
                'البرنامج': program[i],
                'العنوان': title[i],
                'القسم': department[i],
                'رقمقسم': corctnum(int(numdepart[i])),
                'التعليم': edu[i],
                'الطالب': switch_std.get(gen[i]),
                'دكتور': switch_doc.get(supervisorgen[i]),
                'اسمالمشرف': supervisorname[i],
                'تعليمالمشرف': supervisoredu[i],
                'اسمالمشرف2': supervisorname2[i],
                'تعليمالمشرف2': supervisoredu2[i],
                'ذي': switch_to.get(gen[i]),
                'رقمالطلب': ordernumber[i],
                'مدالدراسة': Duration_of_study[i]

            })

        return context
    except Exception:
        return context

