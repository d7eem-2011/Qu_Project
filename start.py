from tkinter import *
from tkinter import filedialog
from tkinter import messagebox as mb
from tkcalendar import DateEntry
from module3 import task_do
import pyarabic.number
from num2words import num2words

# # Most common usage.
# print(num2words(36))
#
# # Other variants, according to the type of article.
# print(num2words(36, to = 'ordinal'))
# print(num2words(36, to = 'ordinal_num'))
# print(num2words(36, to = 'year'))
# print(num2words(36, to = 'currency'))
#
# # Language Support.
# print(num2words(36, lang ='es'))






qux = Tk()

photo = PhotoImage(file=r"C:\Users\D7eem\PycharmProjects\Qu_Project\image\qu1.png")
labelImg = Label(qux, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

qux.title("-الصفحة الرئيسية-")
qux.geometry("600x390")
qux.resizable(False, False)
qux.config( background= "lightblue")


def close():
	qux.destroy()

def oneWindow():
	boneWindow = Toplevel(qux)

	def close():
		boneWindow.destroy()

	def open():
		fileopen = filedialog.askopenfilename(initialdir="/", title="Select A File",
											  filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))

	def save():
		filesave = filedialog.asksaveasfilename(initialdir="/", title="Save A File",
												filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))

	def call():
		res = mb.askquestion('اغلاق النافذة', 'الخروج دون الحفظ')
		if res == 'yes':
			boneWindow.destroy()
		else:
			mb.showinfo('العوده', 'البقاء في النافذة')


	def open_tashkel():
		global fileopen_tashkel
		fileopen_tashkel = filedialog.askopenfilename(initialdir="/", title="Select A File",
													  filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))
		labelFile = Label(boneWindow, text=fileopen_tashkel, background="white")
		labelFile.pack(side=TOP)

	def open_ketat():
		global fileopen_ketat
		fileopen_ketat = filedialog.askopenfilename(
			initialdir="/", title="Select A File", filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))
		labelFile = Label(boneWindow, text=fileopen_ketat, background="white")
		labelFile.pack(side=TOP)


	def task_run():
		docfil = r'C:\Users\D7eem\PycharmProjects\Qu_Project\World Docement\template v4.docx'
		doctash = r"C:\Users\D7eem\PycharmProjects\Qu_Project\World Docement\تشكيل - مذكرات عرض للمجلس.docx"
		try:
			name_clicked = clicked.get()
			name_clicked = 'الجلسة ال'+name_clicked
			task_do(fileopen_ketat, fileopen_tashkel, docfil, doctash, name_clicked)
		except Exception:
			mb.showerror("Warning", "الرجاء اختيار ملف الصحيح !!")

	boneWindow.title("تسجيل جلسة جديدة")
	boneWindow.geometry("600x480")
	boneWindow.resizable(False, False)
	boneWindow.config( background= "lightblue")



	labelImg = Label(boneWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP)
	labelImg.pack(side= TOP)

	l1 = Label(boneWindow, text = "حدد رقم الجلسة", width=15, height=1, background="lightblue", compound=TOP)
	l1.place(x=365, y=180)

	l2 = Label(boneWindow, text="ادخل تاريخ الجلسة", width=15, height=1, background="lightblue", compound=TOP)
	l2.place(x=365, y=220)

	# options = [*range(1,31)]
	#
	# clicked = StringVar()
	# clicked.set(options[0])
	# drop = OptionMenu(boneWindow, clicked, *options)
	# drop.place(x=295, y=175)

	options = []
	for i in range(1,31):
		num = num2words(i, lang='ar', to='ordinal_num')
		options.append(num)

	clicked = StringVar()
	clicked.set(options[0])
	drop = OptionMenu(boneWindow, clicked, *options)
	drop.place(x=155, y=175)

	# cal = DateEntry(boneWindow, width=15, year=2020, month=1, day=1,background='black', foreground='white', borderwidth=2)
	# cal.place(x=240, y=220)

	tashkelButton = Button(boneWindow, text='اختيار ملف التشكيل', width=15, height=1, command=open_tashkel, relief=FLAT)
	tashkelButton.place(x=240, y=260)

	katatButton = Button(boneWindow, text="اختيار ملف الخطط", width=15, height=1, command=open_ketat)
	katatButton.place(x=240, y=300)

	l3 = Label(boneWindow, text=":ملاحظة قبل الحفظ* \n حدد رقم الجلسة - \n حدد تاريخ الجلسة - \n اختيار ملف التشكيل - \n اختيار ملف الخطط -", width=20, height=6, background="#FFFFFF")
	l3.place(x=410, y=300)

	saveButton = Button (boneWindow, text ="حفظ", width=15, height=1, command=task_run)
	saveButton.place(x=240, y=340)

	b4 = Button(boneWindow, text="التحقق من الحفظ", width=15, height=1, command=open)
	b4.place(x=240, y=380)

	buttonClose = Button (boneWindow, text = "اغلاق", width=10, height=1, command = call)
	buttonClose.place(x=255, y=420)

b1 = Button(qux, text="تسجيل جلسة جديدة", width=15, height=1, command=oneWindow).place(x=305, y=180)





def twoWindow():
	btwoWindow = Toplevel(qux)

	def close():
		btwoWindow.destroy()

	def save():
		filesave = filedialog.asksaveasfilename(initialdir="/", title="Save A File",
												filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))

	btwoWindow.title("استعراض الجلسات")
	btwoWindow.geometry("600x400")
	btwoWindow.resizable(False, False)
	btwoWindow.config( background= "lightblue")

	labelImg = Label(btwoWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

	l1 = Label(btwoWindow, text = "حدد رقم الجلسة", width=15, height=1, background="lightblue", compound=TOP).place(x=365, y=180)

	l2 = Label(btwoWindow, text="ادخل تاريخ الجلسة", width=15, height=1, background="lightblue", compound=TOP).place(x=365, y=220)

	options = [ "1", "2", "3","4", "5", "6","7", "8", "9", "10",
				"11", "12", "13","14", "15", "16","17", "18", "19", "20",
				"21", "22", "23","24", "25", "26","27", "28", "29", "30"]

	clicked = StringVar()
	clicked.set(options[0])
	drop = OptionMenu(btwoWindow, clicked, *options)
	drop.place(x=295, y=175)

	options = [ "الاولى", "الثانية", "الثالثة","الرابعة", "الخامسة", "السادسة","السابعة", "الثامنة", "التاسعة", "العاشرة",
				"الحادية عشر", "الثانية عشر", "الثالثة عشر","الرابعة عشر", "الخامسة عشر", "السادسة عشر","السابعة عشر", "الثامنة عشر", "التاسعة عشر", "العشرون",
				"الحادية والعشرون", "الثانية والعشرون", "الثالثة والعشرون","الرابعة والعشرون", "الخامسة والعشرون", "السادسة والعشرون","السابعة والعشرون", "الثامنة والعشرون", "التاسعة والعشرون", "الثلاثون"]

	clicked = StringVar()
	clicked.set(options[0])
	drop = OptionMenu(btwoWindow, clicked, *options)
	drop.place(x=155, y=175)

	cal = DateEntry(btwoWindow, width=15, year=2020, month=1, day=1,
					background='black', foreground='white', borderwidth=2)
	cal.place(x=240, y=220)

	b1 = Button(btwoWindow, text="تنزيل المحضر", width=15, height=1, command=save).place(x=240, y=260)

	b2 = Button(btwoWindow, text="تنزيل اعلان الجلسة", width=15, height=1, command=save).place(x=240, y=300)

	b3 = Button (btwoWindow, text = "اغلاق", width=10, height=1, command = close).place(x=255, y=340)

b2 = Button(qux, text="استعراض الجلسات", width=15, height=1, command=twoWindow).place(x=165, y=180)

def threeWindow():
	bthreeWindow = Toplevel(qux)

	def close():
		bthreeWindow.destroy()

	bthreeWindow.title("التقارير")
	bthreeWindow.geometry("600x400")
	bthreeWindow.resizable(False, False)
	bthreeWindow.config( background= "lightblue")

	labelImg = Label(bthreeWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

	l1 = Label(bthreeWindow,text = ":التاريخ", width=15, height=1, background="lightblue", compound=TOP).place(x=410, y=140)

	cal = DateEntry(bthreeWindow, width=12, year=2020, month=1, day=1,background='black', foreground='white', borderwidth=2)
	cal.place(x=120, y=140)

	l2 = Label(bthreeWindow,text = ":الاسم", width=15, height=1, background="lightblue", compound=TOP)
	l2.place(x=410, y=180)

	t1 = Text(bthreeWindow, width=40, height=1).place(x=120, y=180)

	l3 = Label(bthreeWindow,text = ":الجنس", width=15, height=1, background="lightblue", compound=TOP).place(x=410, y=220)

	options = [ "ذكر", "انثى"]

	clicked = StringVar()
	clicked.set(options[0])
	drop = OptionMenu(bthreeWindow, clicked, *options)
	drop.place(x=120, y=215)

	l4 = Label(bthreeWindow,text = ":التخصص", width=15, height=1, background="lightblue", compound=TOP).place(x=405, y=260)

	t2 = Text(bthreeWindow, width=38, height=1).place(x=120, y=260)

	b1 = Button (bthreeWindow, text = "استعراض", width=15, height=1).place(x=240, y=300)

	b2 = Button (bthreeWindow, text = "اغلاق", width=10, height=1, command = close).place(x=255, y=340)

b3 = Button(qux, text = "التقارير", width=15, height=1, command=threeWindow).place(x=305, y=230)

def fourWindow():
	bfourWindow = Toplevel(qux)

	def call():
		res = mb.askquestion('اغلاق النافذة', 'الخروج دون الحفظ')
		if res == 'yes':
			bfourWindow.destroy()
		else:
			mb.showinfo('العوده', 'البقاء في النافذة')

	bfourWindow.title("اعضاء الجلسة")
	bfourWindow.geometry("600x480")
	bfourWindow.resizable(False, False)
	bfourWindow.config(background="lightblue")

	labelImg = Label(bfourWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

	l1 = Label(bfourWindow,text = ":الاسم", width=15, height=1, background="lightblue", compound=TOP).place(x=410, y=180)

	t1 = Text(bfourWindow, width=40, height=1).place(x=120, y=180)

	l2 = Label(bfourWindow,text = ":الدرجة العلمية", width=15, height=1, background="lightblue", compound=TOP).place(x=390, y=220)

	options = [ "استاذ", "استاذ مساعد", "استاذ مشارك"]

	clicked = StringVar()
	clicked.set(options[0])
	drop = OptionMenu(bfourWindow, clicked, *options)
	drop.place(x=120, y=220)

	l3 = Label(bfourWindow,text = ":الكلية", width=15, height=1, background="lightblue", compound=TOP).place(x=410, y=260)

	t2 = Text(bfourWindow, width=40, height=1).place(x=120, y=260)

	l4 = Label(bfourWindow,text = ":تاريخ التكليف", width=15, height=1, background="lightblue", compound=TOP).place(x=390, y=300)

	cal1 = DateEntry(bfourWindow, width=12, year=2020, month=1, day=1,
					background='black', foreground='white', borderwidth=2)
	cal1.place(x=120, y=300)

	l5 = Label(bfourWindow,text = ":تاريخ انتهاء التكليف", width=15, height=1, background="lightblue", compound=TOP).place(x=375, y=340)

	cal2 = DateEntry(bfourWindow, width=12, year=2020, month=1, day=1,
					background='black', foreground='white', borderwidth=2)
	cal2.place(x=120, y=340)

	b1 = Button(bfourWindow, text = "اضافة عضو", width=15, height=1).place(x=365, y=380)

	b2 = Button(bfourWindow, text = "حفظ", width=15, height=1).place(x=240, y=380)

	b3 = Button(bfourWindow, text="التعديل والحذف", width=15, height=1).place(x=115, y=380)

	b4 = Button(bfourWindow, text = "اغلاق", width=10, height=1, command = call).place(x=255, y=420)

b4 = Button (qux, text = "اعضاء الجلسة", width=15, height=1, command=fourWindow).place(x=165, y=230)

def fiveWindow():
	bfiveWindow = Toplevel(qux)

	def close():
		bfiveWindow.destroy()

	def open1():
		fileopen1 = filedialog.askopenfilename()
		label01 = Label(bfiveWindow, text=fileopen1).place(x=20, y=220)

	def open2():
		fileopen2 = filedialog.askopenfilename()
		label02 = Label(bfiveWindow, text=fileopen2).place(x=20, y=300)

	def open3():
		fileopen3 = filedialog.askopenfilename()
		label03 = Label(bfiveWindow, text=fileopen3).place(x=20, y=380)

	bfiveWindow.title("اعداد البرنامج")
	bfiveWindow.geometry("600x480")
	bfiveWindow.resizable(False, False)
	bfiveWindow.config( background= "lightblue")

	labelImg = Label(bfiveWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

	b1 = Button(bfiveWindow, text="اختيار مسار الملف الاساسي", width=20, height=1, command = open1).place(x=220, y=180)

	b2 = Button(bfiveWindow, text="اختيار مسار الملف الاحتياطي", width=20, height=1, command = open2).place(x=220, y=260)

	b3 = Button(bfiveWindow, text="اختيار مسار ملف الاعضاء", width=20, height=1, command = open3).place(x=220,	y=340)

	b4 = Button(bfiveWindow, text = "اغلاق", width=10, height=1, command = close).place(x=255, y=420)

button5 = Button(qux, text="اعداد البرنامج", width=15, height=1, command=fiveWindow).place(x=165, y=280)

def sixWindow():
	bsixWindow = Toplevel(qux)

	def close():
		bsixWindow.destroy()

	def open():
		fileopen = filedialog.askopenfilename(initialdir="/", title="Select A File",
											  filetypes=(("xlsx", "*.xlsx"), ("all files", "*.*")))

	bsixWindow.title("اعداد الجلسة")
	bsixWindow.geometry("600x480")
	bsixWindow.resizable(False, False)
	bsixWindow.config( background= "lightblue")

	labelImg = Label(bsixWindow, width=377, height=140, image = photo, background= "lightblue", compound = TOP).pack(side= TOP)

	b1 = Button (bsixWindow, text = "تحديد مسار", width=15, height=1, command = open).place(x=240, y=180)

	b2 = Button (bsixWindow, text = "اعداد المحضر", width=15, height=1).place(x=240, y=220)

	b3 = Button (bsixWindow, text = "اعلان الجلسة", width=15, height=1).place(x=240, y=260)

	b4 = Button (bsixWindow, text = "ملخص الاجتماع", width=15, height=1).place(x=240, y=300)

	b5 = Button (bsixWindow, text = "مواضيع الجلسة", width=15, height=1).place(x=240, y=340)

	def attendWindow():
		battendWindow = Toplevel(qux)

		def call():
			res = mb.askquestion('اغلاق النافذة', 'الخروج دون الحفظ')
			if res == 'yes':
				battendWindow.destroy()
			else:
				mb.showinfo('العوده', 'البقاء في النافذة')

		battendWindow.title("اعضاء الجلسة")
		battendWindow.geometry("600x520")
		battendWindow.resizable(False, False)
		battendWindow.config(background="lightblue")

		labelImg = Label(battendWindow, width=377, height=140, image=photo, background="lightblue", compound=TOP)
		labelImg.pack(
			side=TOP)

		l1 = Label(battendWindow, text=":الاسم", width=15, height=1, background="lightblue", compound=TOP).place(x=410,
																											   y=180)

		t1 = Text(battendWindow, width=40, height=1).place(x=120, y=180)

		l2 = Label(battendWindow, text=":الدرجة العلمية", width=15, height=1, background="lightblue", compound=TOP).place(
			x=390, y=220)

		options = ["استاذ", "استاذ مساعد", "استاذ مشارك"]

		clicked = StringVar()
		clicked.set(options[0])
		drop = OptionMenu(battendWindow, clicked, *options)
		drop.place(x=120, y=220)

		l3 = Label(battendWindow, text=":الكلية", width=15, height=1, background="lightblue", compound=TOP).place(x=410,
																												y=260)

		t2 = Text(battendWindow, width=40, height=1)
		t2.place(x=120, y=260)

		l4 = Label(battendWindow, text=":تاريخ التكليف", width=15, height=1, background="lightblue", compound=TOP)
		l4.place(
			x=390, y=300)

		cal1 = DateEntry(battendWindow, width=12, year=2020, month=1, day=1, background='black', foreground='white',
						borderwidth=2)
		cal1.place(x=120, y=300)

		l5 = Label(battendWindow, text=":تاريخ انتهاء التكليف", width=15, height=1, background="lightblue",
				   compound=TOP).place(x=377, y=340)

		cal2 = DateEntry(battendWindow, width=12, year=2020, month=1, day=1,
						 background='black', foreground='white', borderwidth=2)
		cal2.place(x=120, y=340)

		l6 = Label(battendWindow, text=":التحضير", width=15, height=1, background="lightblue", compound=TOP).place(x=405,
																												  y=380)

		var = IntVar()
		R1 = Radiobutton(battendWindow, text="حاضر", variable=var,background="lightblue", value=1).place(x=300, y=380)

		R2 = Radiobutton(battendWindow, text="غائب", variable=var,background="lightblue", value=2).place(x=240, y=380)

		b1 = Button(battendWindow, text="حفظ", width=15, height=1).place(x=240, y=420)

		b2 = Button(battendWindow, text="اغلاق", width=10, height=1, command=call).place(x=255, y=460)

	b6 = Button(bsixWindow, text="تحضير الاعضاء", width=15, height=1, command=attendWindow).place(x=240, y=380)

	b7 = Button(bsixWindow, text = "اغلاق", width=10, height=1, command=close).place(x=255, y=420)

button6 = Button(qux, text="اعداد الجلسة", width=15, height=1, command=sixWindow).place(x=305, y=280)

buttonClose = Button(qux, text = "اغلاق", width=10, height=1, command=close).place(x=255, y=320)

qux.mainloop()