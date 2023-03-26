
from openpyxl import load_workbook
import os
from datetime import date
import cx_Oracle
import tkinter as tk
from tkinter import *
from tkinter import ttk
import pandas as pd


class App:
    def __init__(self,app):
        self.app=app
        self.app.geometry("1920x1080")
        self.app.title("Finance Management System")
        self.co="#2B4257"
        #Label1
        Label1=Label(self.app,bg="#2B4257",fg="white",text="Finance Management System",font="times 15 bold").pack(fill=X)

        #Buttons
        Check=Button(text="Check",bg="yellow",fg="black",font="systemui 15 bold",command=self.check).place(x=650,y=250,width=200,height=40)
        New_Entry = Button(text="New Entry", bg="green", fg="white", font="systemui 15 bold",command=self.new_entry).place(x=650, y=300, width=200,
                                                                                             height=40)
        Close = Button(text="Close", bg="red", fg="white", font="systemui 15 bold",command=self.close).place(x=650, y=350, width=200,
                                                                                             height=40)

    def check(self):
        class EA:
            def __init__(self,app):
                self.app=app
                self.app=self.app
                self.apk=Toplevel(self.app)
                self.apk.geometry("1920x1080")
                self.ea1=StringVar()
                self.ea2=StringVar()
                labelea1=Label(self.apk,text="Fetchig Data",bg="#2B4257",fg="white",font="times 15 bold").pack(fill=X)
                labelea2=Label(self.apk,text="From",bg="skyblue",fg="black",font="systemui 15 bold").place(x=500,y=250,height=50,width=200)
                entryea1=Entry(self.apk,textvariable=self.ea1).place(x=750,y=250,height=50,width=200)
                labelea3=Label(self.apk,text="To",bg="yellow",fg="black",font="systemui 15 bold").place(x=500,y=300,height=50,width=200)
                entryea2=Entry(self.apk,textvariable=self.ea2).place(x=750,y=300,height=50,width=200)
                run=Button(self.apk,text="Run",font="systemui 15 bold",bg="green",fg="white",command=self.run).place(x=600,y=400,height=50,width=150)
                clear = Button(self.apk, text="Clear", font="systemui 15 bold", bg="red", fg="white",command=self.clear).place(x=600, y=450, height=50, width=150)

            def clear(self):
                self.ea1.set("")
                self.ea2.set("")
            def run(self):
                print("hi")
                con = cx_Oracle.connect('database/user123@localhost')
                cursor = con.cursor()

                if self.ea1.get()!="" and self.ea1.get()!="":
                    january = ['01 01', '02 01', '03 01', '04 01', '05 01', '06 01', '07 01', '08 01', '09 01', '10 01',
                               '11 01', '12 01', '13 01', '14 01', '15 01', '16 01', '17 01', '18 01', '19 01', '20 01',
                               '21 01', '22 01', '23 01', '24 01', '25 01', '26 01', '27 01', '28 01', '29 01', '30 01',
                               '31 01']
                    febrauary = ['01 02', '02 02', '03 02', '04 02', '05 02', '06 02', '07 02', '08 02', '09 02',
                                 '10 02', '11 02', '12 02', '13 02', '14 02', '15 02', '16 02', '17 02', '18 02',
                                 '19 02', '20 02', '21 02', '22 02', '23 02', '24 02', '25 02', '26 02', '27 02',
                                 '28 02']
                    march = ['01 03', '02 03', '03 03', '04 03', '05 03', '06 03', '07 03', '08 03', '09 03', '10 03',
                             '11 03', '12 03', '13 03', '14 03', '15 03', '16 03', '17 03', '18 03', '19 03', '20 03',
                             '21 03', '22 03', '23 03', '24 03', '25 03', '26 03', '27 03', '28 03', '29 03', '30 03',
                             '31 03']
                    april = ['01 04', '02 04', '03 04', '04 04', '05 04', '06 04', '07 04', '08 04', '09 04', '10 04',
                             '11 04', '12 04', '13 04', '14 04', '15 04', '16 04', '17 04', '18 04', '19 04', '20 04',
                             '21 04', '22 04', '23 04', '24 04', '25 04', '26 04', '27 04', '28 04', '29 04', '30 04']
                    may = ['01 05', '02 05', '03 05', '04 05', '05 05', '06 05', '07 05', '08 05', '09 05', '10 05',
                           '11 05', '12 05', '13 05', '14 05', '15 05', '16 05', '17 05', '18 05', '19 05', '20 05',
                           '21 05', '22 05', '23 05', '24 05', '25 05', '26 05', '27 05', '28 05', '29 05', '30 05',
                           '31 05']
                    june = ['01 06', '02 06', '03 06', '04 06', '05 06', '06 06', '07 06', '08 06', '09 06', '10 06',
                            '11 06', '12 06', '13 06', '14 06', '15 06', '16 06', '17 06', '18 06', '19 06', '20 06',
                            '21 06', '22 06', '23 06', '24 06', '25 06', '26 06', '27 06', '28 06', '29 06', '30 06']
                    july = ['01 07', '02 07', '03 07', '04 07', '05 07', '06 07', '07 07', '08 07', '09 07', '10 07',
                            '11 07', '12 07', '13 07', '14 07', '15 07', '16 07', '17 07', '18 07', '19 07', '20 07',
                            '21 07', '22 07', '23 07', '24 07', '25 07', '26 07', '27 07', '28 07', '29 07', '30 07',
                            '31 07']
                    august = ['01 08', '02 08', '03 08', '04 08', '05 08', '06 08', '07 08', '08 08', '09 08', '10 08',
                              '11 08', '12 08', '13 08', '14 08', '15 08', '16 08', '17 08', '18 08', '19 08', '20 08',
                              '21 08', '22 08', '23 08', '24 08', '25 08', '26 08', '27 08', '28 08', '29 08', '30 08',
                              '31 08']
                    september = ['01 09', '02 09', '03 09', '04 09', '05 09', '06 09', '07 09', '08 09', '09 09',
                                 '10 09', '11 09', '12 09', '13 09', '14 09', '15 09', '16 09', '17 09', '18 09',
                                 '19 09', '20 09', '21 09', '22 09', '23 09', '24 09', '25 09', '26 09', '27 09',
                                 '28 09', '29 09', '30 09']
                    october = ['01 10', '02 10', '03 10', '04 10', '05 10', '06 10', '07 10', '08 10', '09 10', '10 10',
                               '11 10', '12 10', '13 10', '14 10', '15 10', '16 10', '17 10', '18 10', '19 10', '20 10',
                               '21 10', '22 10', '23 10', '24 10', '25 10', '26 10', '27 10', '28 10', '29 10', '30 10',
                               '31 10']
                    november = ['01 11', '02 11', '03 11', '04 11', '05 11', '06 11', '07 11', '08 11', '09 11',
                                '10 11', '11 11', '12 11', '13 11', '14 11', '15 11', '16 11', '17 11', '18 11',
                                '19 11', '20 11', '21 11', '22 11', '23 11', '24 11', '25 11', '26 11', '27 11',
                                '28 11', '29 11', '30 11']
                    december = ['01 12', '02 12', '03 12', '04 12', '05 12', '06 12', '07 12', '08 12', '09 12',
                                '10 12', '11 12', '12 12', '13 12', '14 12', '15 12', '16 12', '17 12', '18 12',
                                '19 12', '20 12', '21 12', '22 12', '23 12', '24 12', '25 12', '26 12', '27 12',
                                '28 12', '29 12', '30 12', '31 12']
                    year = [january, febrauary, march, april, may, june, july, august, september, october, november,
                            december]
                    ljanuary = ['01 01', '02 01', '03 01', '04 01', '05 01', '06 01', '07 01', '08 01', '09 01',
                                '10 01', '11 01', '12 01', '13 01', '14 01', '15 01', '16 01', '17 01', '18 01',
                                '19 01', '20 01', '21 01', '22 01', '23 01', '24 01', '25 01', '26 01', '27 01',
                                '28 01', '29 01', '30 01', '31 01']
                    lfebrauary = ['01 02', '02 02', '03 02', '04 02', '05 02', '06 02', '07 02', '08 02', '09 02',
                                  '10 02', '11 02', '12 02', '13 02', '14 02', '15 02', '16 02', '17 02', '18 02',
                                  '19 02', '20 02', '21 02', '22 02', '23 02', '24 02', '25 02', '26 02', '27 02',
                                  '28 02', '29 02']
                    lmarch = ['01 03', '02 03', '03 03', '04 03', '05 03', '06 03', '07 03', '08 03', '09 03', '10 03',
                              '11 03', '12 03', '13 03', '14 03', '15 03', '16 03', '17 03', '18 03', '19 03', '20 03',
                              '21 03', '22 03', '23 03', '24 03', '25 03', '26 03', '27 03', '28 03', '29 03', '30 03',
                              '31 03']
                    lapril = ['01 04', '02 04', '03 04', '04 04', '05 04', '06 04', '07 04', '08 04', '09 04', '10 04',
                              '11 04', '12 04', '13 04', '14 04', '15 04', '16 04', '17 04', '18 04', '19 04', '20 04',
                              '21 04', '22 04', '23 04', '24 04', '25 04', '26 04', '27 04', '28 04', '29 04', '30 04']
                    lmay = ['01 05', '02 05', '03 05', '04 05', '05 05', '06 05', '07 05', '08 05', '09 05', '10 05',
                            '11 05', '12 05', '13 05', '14 05', '15 05', '16 05', '17 05', '18 05', '19 05', '20 05',
                            '21 05', '22 05', '23 05', '24 05', '25 05', '26 05', '27 05', '28 05', '29 05', '30 05',
                            '31 05']
                    ljune = ['01 06', '02 06', '03 06', '04 06', '05 06', '06 06', '07 06', '08 06', '09 06', '10 06',
                             '11 06', '12 06', '13 06', '14 06', '15 06', '16 06', '17 06', '18 06', '19 06', '20 06',
                             '21 06', '22 06', '23 06', '24 06', '25 06', '26 06', '27 06', '28 06', '29 06', '30 06']
                    ljuly = ['01 07', '02 07', '03 07', '04 07', '05 07', '06 07', '07 07', '08 07', '09 07', '10 07',
                             '11 07', '12 07', '13 07', '14 07', '15 07', '16 07', '17 07', '18 07', '19 07', '20 07',
                             '21 07', '22 07', '23 07', '24 07', '25 07', '26 07', '27 07', '28 07', '29 07', '30 07',
                             '31 07']
                    laugust = ['01 08', '02 08', '03 08', '04 08', '05 08', '06 08', '07 08', '08 08', '09 08', '10 08',
                               '11 08', '12 08', '13 08', '14 08', '15 08', '16 08', '17 08', '18 08', '19 08', '20 08',
                               '21 08', '22 08', '23 08', '24 08', '25 08', '26 08', '27 08', '28 08', '29 08', '30 08',
                               '31 08']
                    lseptember = ['01 09', '02 09', '03 09', '04 09', '05 09', '06 09', '07 09', '08 09', '09 09',
                                  '10 09', '11 09', '12 09', '13 09', '14 09', '15 09', '16 09', '17 09', '18 09',
                                  '19 09', '20 09', '21 09', '22 09', '23 09', '24 09', '25 09', '26 09', '27 09',
                                  '28 09', '29 09', '30 09']
                    loctober = ['01 10', '02 10', '03 10', '04 10', '05 10', '06 10', '07 10', '08 10', '09 10',
                                '10 10', '11 10', '12 10', '13 10', '14 10', '15 10', '16 10', '17 10', '18 10',
                                '19 10', '20 10', '21 10', '22 10', '23 10', '24 10', '25 10', '26 10', '27 10',
                                '28 10', '29 10', '30 10', '31 10']
                    lnovember = ['01 11', '02 11', '03 11', '04 11', '05 11', '06 11', '07 11', '08 11', '09 11',
                                 '10 11', '11 11', '12 11', '13 11', '14 11', '15 11', '16 11', '17 11', '18 11',
                                 '19 11', '20 11', '21 11', '22 11', '23 11', '24 11', '25 11', '26 11', '27 11',
                                 '28 11', '29 11', '30 11']
                    ldecember = ['01 12', '02 12', '03 12', '04 12', '05 12', '06 12', '07 12', '08 12', '09 12',
                                 '10 12', '11 12', '12 12', '13 12', '14 12', '15 12', '16 12', '17 12', '18 12',
                                 '19 12', '20 12', '21 12', '22 12', '23 12', '24 12', '25 12', '26 12', '27 12',
                                 '28 12', '29 12', '30 12', '31 12']
                    leap_year = [ljanuary, lfebrauary, lmarch, lapril, lmay, ljune, ljuly, laugust, lseptember,
                                 loctober, lnovember, ldecember]
                    y = []
                    for x in year:
                        for j in x:
                            y.append(j)
                    ly = []
                    for w in leap_year:
                        for e in w:
                            ly.append(e)

                    s = self.ea1.get()
                    l = self.ea2.get()
                    q = 0
                    z = 0
                    en = 0
                    ex = 0
                    uu = []
                    if int(s[6] + s[7]) == int(l[6] + l[7]):
                        if int(s[6] + s[7]) % 4 == 0:
                            for i in y:
                                if str(i) == str(s[0] + s[1] + " " + s[3] + s[4]):
                                    q = y.index(i)

                            for p in y:
                                if str(p) == str(l[0] + l[1] + " " + l[3] + l[4]):
                                    z = y.index(p)
                            ss = y[q:z + 1]
                            for b in ss:
                                uu.append(b + " " + str(s[6] + s[7]))

                        else:
                            for i in ly:
                                if str(i) == str(s[0] + s[1] + " " + s[3] + s[4]):
                                    q = ly.index(i)

                            for p in ly:
                                if str(p) == str(l[0] + l[1] + " " + l[3] + l[4]):
                                    z = ly.index(p)
                            ss = ly[q:z + 1]
                            for b in ss:
                                uu.append(b + " " + str(s[6] + s[7]))
                    else:
                        if int(l[6] + l[7]) - int(s[6] + s[7]) == 1:
                            if int(s[6] + s[7]) % 4 == 0:
                                for i in ly:
                                    if str(i) == str(s[0] + s[1] + " " + s[3] + s[4]):
                                        q = ly.index(i)
                                ss = ly[q::]
                                for b in ss:
                                    uu.append(b + " " + str(s[6] + s[7]))
                                for ii in y:
                                    if str(ii) == str(l[0] + l[1] + " " + l[3] + l[4]):
                                        q = y.index(ii)
                                ss1 = y[0:q + 1]

                                for bb in ss1:
                                    uu.append(bb + " " + str(l[6] + l[7]))
                            if int(l[6] + l[7]) % 4 == 0:
                                for i in y:
                                    if str(i) == str(s[0] + s[1] + " " + s[3] + s[4]):
                                        q = y.index(i)
                                ss = y[q::]
                                for b in ss:
                                    uu.append(b + " " + str(s[6] + s[7]))
                                for ii in ly:
                                    if str(ii) == str(l[0] + l[1] + " " + l[3] + l[4]):
                                        q = ly.index(ii)
                                ss1 = ly[0:q + 1]
                                for bb in ss1:
                                    uu.append(bb + " " + str(l[6] + l[7]))
                    sql = "SELECT * FROM data2 WHERE pdate IN "

                    final = tuple(uu)
                    final1 = str(final)
                    cursor.execute(sql + final1)
                    self.row = cursor.fetchall()
                if self.ea1.get()!="" and self.ea2.get()=="":
                    sql="SELECT * FROM data2 WHERE pdate"+str(self.ea1.get())
                    cursor.execute(sql)
                    self.row=cursor.fetchall()

                self.check1=Toplevel(self.apk)
                self.check1.geometry("600x300")
                self.check1.title("Details")
                btnp=Button(self.check1,fg="white",bg="green",text="Save",command=self.save).place(x=400,y=250)

                self.total=0

                for tot in self.row:
                    self.total=self.total+int(tot[2])

                labeltot=Label(self.check1,fg="white",bg="purple",text=str(self.total)).place(x=350,y=250)
                tree=ttk.Treeview(self.check1)
                tree['show']='headings'

                sss=ttk.Style(self.check1)
                sss.theme_use("clam")

                sss.configure(".", font=('Helvetica',11))
                sss.configure("Treeview.Heading", foreground='red',font=('Helvetica',11,"bold"))
                tree["columns"]=("date","name","price")
                tree.column("date",width=50,minwidth=50,anchor=tk.CENTER)
                tree.column("name",width=100,minwidth=100,anchor=tk.CENTER)
                tree.column("price",width=50,minwidth=50,anchor=tk.CENTER)
                tree.heading("date",text="date",anchor=tk.CENTER)
                tree.heading("name",text="name",anchor=tk.CENTER)
                tree.heading("price",text="price",anchor=tk.CENTER)

                m=0

                for ro in self.row:
                    tree.insert("",m,text="",values=(ro[0],ro[1],ro[2]))
                    m=m+1

                hsb=ttk.Scrollbar(self.check1,orient="vertical")

                hsb.configure(command=tree.yview)
                tree.configure(yscrollcommand=hsb.set)
                hsb.pack(fill=Y,side=RIGHT)
                tree.pack()
            def save(self):
                all_date=[]
                all_name=[]
                all_price=[]
                for p_date,p_name,p_price in self.row:
                    all_date.append(p_date)
                    all_name.append(p_name)
                    all_price.append(p_price)

                dic ={'p_date':all_date,'p_name':all_name,'p_price':all_price}
                df=pd.DataFrame (dic)
                rrr=str(self.ea1.get()[0])+str(self.ea1.get()[1])+" "+str(self.ea1.get()[3])+str(self.ea1.get()[4])+" "+str(self.ea1.get()[6])+str(self.ea1.get()[7])+"to"+str(self.ea2.get()[0])+str(self.ea2.get()[1])+" "+str(self.ea2.get()[3])+str(self.ea2.get()[4])+" "+str(self.ea2.get()[6])+str(self.ea2.get()[7])
                df_csv=df.to_csv('excel statement/'+rrr+"Account"+".csv")



        EA(self.app)

    def new_entry(self):
        class EE:
            def __init__(self,app):
                self.app=app
                self.app=self.app
                self.root = Toplevel(self.app)
                self.root.geometry("1920x1080")
                self.root.title("Personel Finance Management System")
                self.co = "#2B4257"

                self.a1 = StringVar()
                self.a2 = StringVar()
                self.a3 = StringVar()
                self.a4 = StringVar()
                self.a5 = StringVar()
                self.a6 = StringVar()
                self.a7 = StringVar()
                self.a8 = StringVar()
                self.a9 = StringVar()
                self.a10 = StringVar()
                self.a11 = StringVar()

                self.b1 = IntVar()
                self.b2 = IntVar()
                self.b3 = IntVar()
                self.b4 = IntVar()
                self.b5 = IntVar()
                self.b6 = IntVar()
                self.b7 = IntVar()
                self.b8 = IntVar()
                self.b9 = IntVar()
                self.b10 = IntVar()
                self.b11 = IntVar()

                self.d = StringVar()
                self.ex= StringVar()

                label1 = Label(self.root, text="Personal Finance Management System", font="fiver 15 bold", bg=self.co,
                               fg="white").pack(fill=X)

                self.Frame2 = Frame(self.root, bg=self.co)
                self.Frame2.place(x=100, y=100, height=500, width=900)
                self.Frame3 = Frame(self.Frame2, bg="grey")
                self.Frame3.place(height=30, width=900)
                label1a = Label(self.Frame3, text="Item you Purchased", bg="grey", fg="white").pack(fill=X)
                self.Frame4 = Frame(self.Frame2, bg=self.co)
                self.Frame4.place(y=35)

                label2 = Label(self.Frame4, text="Name", bg="#555", fg="white", font="systemui 15 bold", padx=50).grid(row=1,
                                                                                                                       column=1)
                label3 = Label(self.Frame4, text="Amount", bg="#555", fg="white", font="systemui 15 bold").grid(row=1, column=2,
                                                                                                                padx=30)
                entry1a = Entry(self.Frame4, textvariable=self.a1).grid(row=2, column=1, padx=20, pady=10)
                entry2a = Entry(self.Frame4, textvariable=self.a2).grid(row=3, column=1, padx=20, pady=10)
                entry3a = Entry(self.Frame4, textvariable=self.a3).grid(row=4, column=1, padx=20, pady=10)
                entry4a = Entry(self.Frame4, textvariable=self.a4).grid(row=5, column=1, padx=20, pady=10)
                entry5a = Entry(self.Frame4, textvariable=self.a5).grid(row=6, column=1, padx=20, pady=10)
                entry6a = Entry(self.Frame4, textvariable=self.a6).grid(row=7, column=1, padx=20, pady=10)
                entry7a = Entry(self.Frame4, textvariable=self.a7).grid(row=8, column=1, padx=20, pady=10)
                entry8a = Entry(self.Frame4, textvariable=self.a8).grid(row=9, column=1, padx=20, pady=10)
                entry9a = Entry(self.Frame4, textvariable=self.a9).grid(row=10, column=1, padx=20, pady=10)
                entry10a = Entry(self.Frame4, textvariable=self.a10).grid(row=11, column=1, padx=20, pady=10)
                entry11a = Entry(self.Frame4, textvariable=self.a11).grid(row=12, column=1, padx=20, pady=10)
                entry1b = Entry(self.Frame4, textvariable=self.b1).grid(row=2, column=2, padx=20, pady=10)
                entry2b = Entry(self.Frame4, textvariable=self.b2).grid(row=3, column=2, padx=20, pady=10)
                entry3b = Entry(self.Frame4, textvariable=self.b3).grid(row=4, column=2, padx=20, pady=10)
                entry4b = Entry(self.Frame4, textvariable=self.b4).grid(row=5, column=2, padx=20, pady=10)
                entry5b = Entry(self.Frame4, textvariable=self.b5).grid(row=6, column=2, padx=20, pady=10)
                entry6b = Entry(self.Frame4, textvariable=self.b6).grid(row=7, column=2, padx=20, pady=10)
                entry7b = Entry(self.Frame4, textvariable=self.b7).grid(row=8, column=2, padx=20, pady=10)
                entry8b = Entry(self.Frame4, textvariable=self.b8).grid(row=9, column=2, padx=20, pady=10)
                entry9b = Entry(self.Frame4, textvariable=self.b9).grid(row=10, column=2, padx=20, pady=10)
                entry10b = Entry(self.Frame4, textvariable=self.b10).grid(row=11, column=2, padx=20, pady=10)
                entry11b = Entry(self.Frame4, textvariable=self.b11).grid(row=12, column=2, padx=20, pady=10)
                labeld = Label(self.Frame4, text="Date", fg="white", bg=self.co).grid(row=4, column=3)
                entryd = Entry(self.Frame4, textvariable=self.d).grid(row=4, column=4)
                Btn1 = Button(self.Frame4, fg="white", bg="green", text="Submit", command=self.sub).grid(row=3, column=3)
                Btn2 = Button(self.Frame4, fg="white", bg="red", text="Clear",command=self.clear).grid(row=3, column=4)
                entryex=Entry(self.Frame4,textvariable=self.ex).grid(row=5,column=5)

            def sub(self):
                wb = load_workbook(r'excel/spreadsheet.xlsx')
                sheet = wb.active


                sheet['E10'] = str(self.a1.get())
                sheet['E11'] = str(self.a2.get())
                sheet['E12'] = str(self.a3.get())
                sheet['E13'] = str(self.a4.get())
                sheet['E14'] = str(self.a5.get())
                sheet['E15'] = str(self.a6.get())
                sheet['E16'] = str(self.a7.get())
                sheet['E17'] = str(self.a8.get())
                sheet['E18'] = str(self.a9.get())
                sheet['E19'] = str(self.a10.get())
                sheet['E20'] = str(self.a11.get())

                sheet['I10'] = int(self.b1.get())
                sheet['I11'] = int(self.b2.get())
                sheet['I12'] = int(self.b3.get())
                sheet['I13'] = int(self.b4.get())
                sheet['I14'] = int(self.b5.get())
                sheet['I15'] = int(self.b6.get())
                sheet['I16'] = int(self.b7.get())
                sheet['I17'] = int(self.b8.get())
                sheet['I18'] = int(self.b9.get())
                sheet['I19'] = int(self.b10.get())
                sheet['I20'] = int(self.b11.get())

                if str(self.d.get())!="":
                    self.q = "excel/" + str(self.d.get()[0]) + str(self.d.get()[1]) + str(self.d.get()[3]) + str(self.d.get()[4]) + str(self.d.get()[6]) + str(self.d.get()[7]) + " " + " Account" +self.ex.get()+ ".xlsx"
                    wb.save(self.q)
                    sheet['F3']=str(self.d.get())
                    con = cx_Oracle.connect('database/user123@localhost')
                    cursor = con.cursor()
                    dates=str(self.d.get()[0]) + str(self.d.get()[1]) +" "+ str(self.d.get()[3]) + str(
                        self.d.get()[4]) +" "+ str(self.d.get()[6]) + str(self.d.get()[7])
                    if self.a1.get()!="" and self.b1.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a1.get()) + "'" + "," + str(self.b1.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a2.get()!="" and self.b2.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a2.get()) + "'" + "," + str(self.b2.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a3.get()!="" and self.b3.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a3.get()) + "'" + "," + str(self.b3.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a4.get()!="" and self.b4.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a4.get()) + "'" + "," + str(self.b4.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a5.get()!="" and self.b5.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a5.get()) + "'" + "," + str(self.b5.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a6.get()!="" and self.b6.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a6.get()) + "'" + "," + str(self.b6.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a7.get()!="" and self.b7.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a7.get()) + "'" + "," + str(self.b7.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a8.get()!="" and self.b8.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a8.get()) + "'" + "," + str(self.b8.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a9.get()!="" and self.b9.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a9.get()) + "'" + "," + str(self.b9.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a10.get()!="" and self.b10.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(
                            self.a10.get()) + "'" + "," + str(self.b10.get()) + ")"
                        cursor.execute(sql)
                        con.commit()
                    if self.a11.get()!="" and self.b11.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a11.get()) + "'" + "," + str(self.b11.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                else:
                    c = date.today()
                    today=str(c)
                    sheet['F3'] = str(today[8])+str(today[9])+" "+str(today[5])+str(today[6])+" "+str(today[2])+str(today[3])
                    self.q="excel/"+str(today[8])+str(today[9])+str(today[5])+str(today[6])+str(today[2])+str(today[3])+" "+"Account"+self.ex.get()+".xlsx"
                    wb.save(self.q)
                    con = cx_Oracle.connect('database/user123@localhost')
                    cursor = con.cursor()
                    dates=str(today[8])+str(today[9])+" "+str(today[5])+str(today[6])+" "+str(today[2])+str(today[3])
                    if self.a1.get()!="" and self.b1.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a1.get()) + "'" + "," + str(self.b1.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a2.get()!="" and self.b2.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a2.get()) + "'" + "," + str(self.b2.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a3.get()!="" and self.b3.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a3.get()) + "'" + "," + str(self.b3.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a4.get()!="" and self.b4.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(
                            self.a4.get()) + "'" + "," + str(self.b4.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a5.get()!="" and self.b5.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a5.get()) + "'" + "," + str(self.b5.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a6.get()!="" and self.b6.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a6.get()) + "'" + "," + str(self.b6.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a7.get()!="" and self.b7.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a7.get()) + "'" + "," + str(self.b7.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a8.get()!="" and self.b8.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a8.get()) + "'" + "," + str(self.b8.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a9.get()!="" and self.b9.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a9.get()) + "'" + "," + str(self.b9.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a10.get()!="" and self.b10.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a10.get()) + "'" + "," + str(self.b10.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

                    if self.a11.get()!="" and self.b11.get()!=0:
                        sql = "INSERT INTO data2 VALUES(" + "'" + dates + "'" + "," + "'" + str(self.a11.get()) + "'" + "," + str(self.b11.get()) + ")"
                        cursor.execute(sql)
                        con.commit()

            def clear(self):
                self.a1.set("")
                self.a2.set("")
                self.a3.set("")
                self.a4.set("")
                self.a5.set("")
                self.a6.set("")
                self.a7.set("")

                self.a8.set("")
                self.a9.set("")
                self.a10.set("")
                self.a11.set("")

                self.b1.set(0)
                self.b2.set(0)
                self.b3.set(0)
                self.b4.set(0)
                self.b5.set(0)
                self.b6.set(0)
                self.b7.set(0)
                self.b8.set(0)
                self.b9.set(0)
                self.b10.set(0)
                self.b11.set(0)

                self.d.set("")

        EE(self.app)
    def close(self):
        self.app.destroy()
app=Tk()
App(app)
app.mainloop()