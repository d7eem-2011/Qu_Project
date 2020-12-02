import sqlite3


conn = sqlite3.connect('database.db')

c = conn.cursor()


c.execute(""" CREATE TABLE IF NOT EXISTS members(
        name text,
        degre text,
        collge text,
        date_assine text,
        date_of_end_assine text)
        """)

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
