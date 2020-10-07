#load libraries
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


docfil = r'C:\Users\D7eem\PycharmProjects\Qu_Project\World Docement\template v4.docx'

xlxfil = r'C:\Users\D7eem\Desktop\OneDrive_3_10-5-2020\خطط الجلسة الثالثة.xlsx'

doc = DocxTemplate(docfil)


def ketat_def(xlxfil):

    # load .docx and .xlsx filesf

    ex = pd.read_excel(xlxfil, header=0)

    # Replacing any empty cell with 'لايوجد'
    # ex = ex.fillna(0)

    stdname = pd.Series(ex['اسم الطالب'])
    # ex['الرقم الجامعي']= ex['الرقم الجامعي'].astype(int)
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


    #
    # quoat = pd.Series(ex['نسبة الإقتباس'])
    # disc1 = pd.Series(ex['اسم المناقش (1)'])
    # rank1 = pd.Series(ex['الرتبة العلمية للمناقش (1)'])
    # cole1 = pd.Series(ex['القسم الذي يتبعه المناقش (1)'])
    # work1 = pd.Series(ex['جهة عمل المناقش (1)'])
    # adje1 = pd.Series(ex['صفة المناقش (1)'])
    #
    # disc2 = pd.Series(ex['اسم المناقش (2)'])
    # rank2 = pd.Series(ex['الرتبة العلمية للمناقش (2)'])
    # cole2 = pd.Series(ex['القسم الذي يتبعه المناقش (2)'])
    # work2 = pd.Series(ex['جهة عمل المناقش (2)'])
    # adje2 = pd.Series(ex['صفة المناقش (2)'])
    #
    # disc3 = pd.Series(ex['اسم المناقش (3)'])
    # rank3 = pd.Series(ex['الرتبة العلمية للمناقش (3)'])
    # cole3 = pd.Series(ex['القسم الذي يتبعه المناقش (3)'])
    # work3 = pd.Series(ex['جهة عمل المناقش (3)'])
    # adje3 = pd.Series(ex['صفة المناقش (3)'])
    #
    # disc6 = pd.Series(ex['اسم المناقش (6)'])
    # rank6 = pd.Series(ex['الرتبة العلمية للمناقش (6)'])
    # cole6 = pd.Series(ex['القسم الذي يتبعه المناقش (6)'])
    # work6 = pd.Series(ex['جهة عمل المناقش (6)'])
    # adje6 = pd.Series(ex['صفة المناقش (6)'])
    #
    # disc7 = pd.Series(ex['اسم المناقش (7)'])
    # rank7 = pd.Series(ex['الرتبة العلمية للمناقش (7)'])
    # cole7 = pd.Series(ex['القسم الذي يتبعه المناقش (7)'])
    # work7 = pd.Series(ex['جهة عمل المناقش (7)'])
    # adje7 = pd.Series(ex['صفة المناقش (7)'])




    total = stdnumber.count()


    context  = {'خطط': []}

    switch_std = {
        'ذكر': 'طالب',
        'أنثى': 'طالبة'
    }

    switch_doc = {
        'ذكر': 'طالب',
        'أنثى': 'طالبة'
    }

    for i in range(total):
        context['خطط'].append({

        'رقم': pyarabic.number.number2ordinal(i+1),
        'اسمالطالب': stdname[i],
        'رقمالطالب': stdnumber[i],
        'الكلية': cole[i],
        'تاريخالقسم': departdate[i],
        'تاريخالكلية': coledate[i],
        'رقمجلسكلية': pyarabic.number.number2ordinal(numcole[i].astype(int)),
        'البرنامج': program[i],
        'العنوان': title[i],
        'القسم': department[i],
        'رقمقسم': pyarabic.number.number2ordinal(numdepart[i]),
        'التعليم': edu[i],
        'الطالب': switch_std.get(gen[i]),
        'دكتور': switch_doc.get(gen[i]),
        'اسمالمشرف': supervisorname[i],
        'تعليمالمشرف': supervisoredu[i],
        'اسمالمشرف2': supervisorname2[i],
        'تعليمالمشرف2': supervisoredu2[i]

        # 'اقتباس': quoat[i],
        # 'مناقش1': disc1[i],
        # 'رتبة1': rank1[i],
        # 'قسممناقش1': cole1[i],
        # 'عملمناقش1': work1[i],
        # 'صفة1': adje1[i],
        #
        # 'مناقش2': disc2[i],
        # 'رتبة2': rank2[i],
        # 'صفة2': adje2[i],
        # 'قسممناقش2': cole2[i],
        # 'عملمناقش2': work2[i],
        #
        # 'مناقش3': disc3[i],
        # 'رتبة3': rank3[i],
        # 'صفة3': adje3[i],
        # 'قسممناقش3': cole3[i],
        # 'عملمناقش3': work3[i],
        #
        # 'مناقش6': disc6[i],
        # 'رتبة6': rank6[i],
        # 'صفة6': adje6[i],
        # 'قسممناقش6': cole6[i],
        # 'عملمناقش6': work6[i],
        #
        # 'مناقش7': disc7[i],
        # 'رتبة7': rank7[i],
        # 'صفة7': adje7[i],
        # 'قسممناقش7': cole7[i],
        # 'عملمناقش7': work7[i],



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
