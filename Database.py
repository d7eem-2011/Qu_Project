#!/usr/bin/python
# -*- coding=utf-8 -*-
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

    # c.execute(""" CREATE TABLE IF NOT EXISTS tashkel( اسم الطالب text,
    #     الرقم الجامعي int,
    #     الجنس  text,
    #     الكلية  text,
    #     القسم text,
    #     المرحلة (ماجستير - دكتوراه) text,
    #     البرنامج (التخصص) text,
    #     رقم جلسة القسم int,
    #     تاريخ جلسة القسم text,
    #     رقم جلسة الكلية int,
    #     تاريخ جلسة الكلية text,
    #     نسبة الإقتباس int,
    #     عنوان الرسالة text,
    #     رقم المعاملة int
    #     )""")

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

def viewData():
    conn = sqlite3.connect('database.db')
    cur =conn.cursor()
    cur.execute("SELECT * FROM members")
    rows = cur.fetchall()
    conn.close()
    return rows


createtabel()
