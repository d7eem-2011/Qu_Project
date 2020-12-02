import sqlite3
import pandas as pd
import sqlalchemy


def createtabel():
    conn = sqlite3.connect('database.db')

    c = conn.cursor()


    c.execute(""" CREATE TABLE IF NOT EXISTS members(
            name text,
            degre text,
            collge text,
            date_assine text,
            date_of_end_assine text)
            """)

    c.execute(""" CREATE TABLE IF NOT EXISTS tashkel(
                اسم الطالب text,
                الرقم الجامعي int,
                الجنس  text,
                الكلية  text,
                القسم text,
                المرحلة (ماجستير - دكتوراه) text,
                البرنامج (التخصص) text,
                رقم جلسة القسم int,
                تاريخ جلسة القسم text,
                رقم جلسة الكلية int,
                تاريخ جلسة الكلية text,
                نسبة الإقتباس int,
                عنوان الرسالة text,
                رقم المناقش (1) int,
                اسم المناقش (1) text,
                الرتبة العلمية للمناقش (1) text,
                القسم الذي يتبعه المناقش (1) text,
                جهة عمل المناقش (1) text,
                صفة المناقش (1) text,
                رقم المناقش (1) int,
                اسم المناقش (2) text,
                الرتبة العلمية للمناقش (2) text,
                القسم الذي يتبعه المناقش (2) text,
                جهة عمل المناقش (2) text,
                صفة المناقش (2) text,
                رقم المناقش (2) int,
                اسم المناقش (3) text,
                الرتبة العلمية للمناقش (3) text,
                القسم الذي يتبعه المناقش (3) text,
                جهة عمل المناقش (3) text,
                صفة المناقش (3) text,
                رقم المناقش (4) int,
                اسم المناقش (4) text,
                الرتبة العلمية للمناقش (4) text,
                القسم الذي يتبعه المناقش (4) text,
                جهة عمل المناقش (4) text,
                صفة المناقش (4) text,
                رقم المناقش (5) int,
                اسم المناقش (5) text,
                الرتبة العلمية للمناقش (5) text,
                القسم الذي يتبعه المناقش (5) text,
                جهة عمل المناقش (5) text,
                صفة المناقش (5) text,
                رقم المناقش (6) int,
                اسم المناقش (6) text,
                الرتبة العلمية للمناقش (6) text,
                القسم الذي يتبعه المناقش (6) text,
                جهة عمل المناقش (6) text,
                صفة المناقش (6) text,
                رقم المناقش (7) int,
                اسم المناقش (7) text,
                الرتبة العلمية للمناقش (7) text,
                القسم الذي يتبعه المناقش (7) text,
                جهة عمل المناقش (7) text,
                صفة المناقش (7) text,
                جلسة القسم كتابة text,
                جلسة الكلية كتابة text,
                رقم المعاملة int
                
                
                 )""")

    conn.commit()

    conn.close()


def insert_data(name, degre, collge, date_assine, date_assine_end):
    conn = sqlite3.connect('database.db')
    c = conn.cursor()

    c.execute(" INSERT INTO members VALUES ( :name, :degre, :collge, :date_assine, :date_assine_end)",

{
        "name": name,
        "degre": degre,
        "collge": collge,
        "date_assine": date_assine,
        "date_assine_end": date_assine_end 
}
)

    conn.commit()
    conn.close()


def tashkelcope(excal):
    engine = sqlalchemy.create_engine('sqlite:///database.db')
    df = pd.read_excel(excal)

createtabel()