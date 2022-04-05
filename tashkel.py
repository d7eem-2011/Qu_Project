
import time
import pyarabic.number as nb

import pandas as pd





v = time.time()
t = time.ctime(v)



def tashkel_def(X):
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
    quoat = pd.Series(ex['نسبة الإقتباس'])
    title = pd.Series(ex['عنوان الرسالة'])

    disc1 = pd.Series(ex['اسم المناقش (1)'])
    rank1 = pd.Series(ex['الرتبة العلمية للمناقش (1)'])
    cole1 = pd.Series(ex['القسم الذي يتبعه المناقش (1)'])
    work1 = pd.Series(ex['جهة عمل المناقش (1)'])
    adje1 = pd.Series(ex['صفة المناقش (1)'])

    disc2 = pd.Series(ex['اسم المناقش (2)'])
    rank2 = pd.Series(ex['الرتبة العلمية للمناقش (2)'])
    cole2 = pd.Series(ex['القسم الذي يتبعه المناقش (2)'])
    work2 = pd.Series(ex['جهة عمل المناقش (2)'])
    adje2 = pd.Series(ex['صفة المناقش (2)'])

    disc3 = pd.Series(ex['اسم المناقش (3)'])
    rank3 = pd.Series(ex['الرتبة العلمية للمناقش (3)'])
    cole3 = pd.Series(ex['القسم الذي يتبعه المناقش (3)'])
    work3 = pd.Series(ex['جهة عمل المناقش (3)'])
    adje3 = pd.Series(ex['صفة المناقش (3)'])

    disc4 = pd.Series(ex['اسم المناقش (4)'])
    rank4 = pd.Series(ex['الرتبة العلمية للمناقش (4)'])
    cole4 = pd.Series(ex['القسم الذي يتبعه المناقش (4)'])
    work4 = pd.Series(ex['جهة عمل المناقش (4)'])
    adje4 = pd.Series(ex['صفة المناقش (4)'])

    disc5 = pd.Series(ex['اسم المناقش (5)'])
    rank5 = pd.Series(ex['الرتبة العلمية للمناقش (5)'])
    cole5 = pd.Series(ex['القسم الذي يتبعه المناقش (5)'])
    work5 = pd.Series(ex['جهة عمل المناقش (5)'])
    adje5 = pd.Series(ex['صفة المناقش (5)'])

    disc6 = pd.Series(ex['اسم المناقش (6)'])
    rank6 = pd.Series(ex['الرتبة العلمية للمناقش (6)'])
    cole6 = pd.Series(ex['القسم الذي يتبعه المناقش (6)'])
    work6 = pd.Series(ex['جهة عمل المناقش (6)'])
    adje6 = pd.Series(ex['صفة المناقش (6)'])

    disc7 = pd.Series(ex['اسم المناقش (7)'])
    rank7 = pd.Series(ex['الرتبة العلمية للمناقش (7)'])
    cole7 = pd.Series(ex['القسم الذي يتبعه المناقش (7)'])
    work7 = pd.Series(ex['جهة عمل المناقش (7)'])
    adje7 = pd.Series(ex['صفة المناقش (7)'])
    Duration_of_study = pd.Series(ex['مدة الدراسة'])

    Transactionnumber = pd.Series(ex['رقم المعاملة'])

    total = stdnumber.count()

    context = {'تشكيل': []}

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
    def change(x):
        if type(x) == str:
            word ='القصيم'
            if word in x:
                return "مناقش داخلي"
            else:
                return "مناقش خارجي"
        else:
            return 'لايوجد'

    def corctnum(x, y=True):
        if x + 1 == 1:
            return 'الاول'
        else:
            return nb.number2ordinal(x + 1, feminin=y)

    for i in range(total):
        if stdname[i] == 0 or stdnumber[i] == 0:
            continue

        context['تشكيل'].append({
            'رقم': corctnum(i, y=False),
            'اسمالطالب': stdname[i],
            'رقمالطالب': stdnumber[i],
            'الكلية': cole[i],
            'القسم': department[i],
            'التعليم': edu[i],
            'البرنامج': program[i],
            'رقمقسم': corctnum(int(numdepart[i])),
            'تاريخالقسم': departdate[i],
            'رقمجلسكلية': corctnum(int(numcole[i])),
            'تاريخالكلية': coledate[i],
            'العنوان': title[i],
            'اقتباس': quoat[i],
            'الطالب': switch_std.get(gen[i]),
            'دكتور': switch_doc.get(gen[i]),

            'مناقش1': disc1[i],
            'رتبة1': rank1[i],
            'قسممناقش1': cole1[i],
            'عملمناقش1': work1[i],
            'صفة1': adje1[i],

            'مناقش2': disc2[i],
            'رتبة2': rank2[i],
            'صفة2': change(work2[i]),
            'قسممناقش2': cole2[i],
            'عملمناقش2': work2[i],

            'مناقش3': disc3[i],
            'رتبة3': rank3[i],
            'صفة3': change(work3[i]),
            'قسممناقش3': cole3[i],
            'عملمناقش3': work3[i],

            'مناقش4': disc4[i],
            'رتبة4': rank4[i],
            'صفة4': change(work4[i]),
            'قسممناقش4': cole4[i],
            'عملمناقش4': work4[i],

            'مناقش5': disc5[i],
            'رتبة5': rank5[i],
            'صفة5': change(work5[i]),
            'قسممناقش5': cole5[i],
            'عملمناقش5': work5[i],

            'مناقش6': disc6[i],
            'رتبة6': rank6[i],
            'صفة6': change(work6[i]),
            'قسممناقش6': cole6[i],
            'عملمناقش6': work6[i],

            'مناقش7': disc7[i],
            'رتبة7': rank7[i],
            'صفة7': change(work7[i]),
            'قسممناقش7': cole7[i],
            'عملمناقش7': work7[i],
            'ذي': switch_to.get(gen[i]),
            'مدالدراسة': Duration_of_study[i],
            'رقمالمعاملة': Transactionnumber[i]

        })
    return context
