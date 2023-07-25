from tkinter import *
from tkinter import messagebox
from datetime import date
from tkinter import filedialog
from PIL import Image,ImageTk

import os
from tkinter.ttk import Combobox
import openpyxl, xlrd
from openpyxl import Workbook
import pathlib
import tkinter as tk



window= Tk()

window.title("Student Management")
window.geometry("1250x700")
window.config(background="white")

search_image= PhotoImage(file="small_Search_icon.png")
update_image= PhotoImage(file="small_change_icon.png")
#
# def select_gender():
#     value=gender.get()
#     if value == 1:
#         gender = "Male"
#         print(gender)
#     elif value == 2:
#         gender = "Female"
#         print(gender)
#     elif value == 3:
#         gender = "Other"
#         print(gender)
#     else:
#         gender="Prefer Not To Say"
#
#
def Upload_photo():
    filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                          title="Select an Image file",
                                          filetype=(("JPG file", "*.jpg"),
                                                    ("PNG file", "*.png"),
                                                    ("All files", "*.*")))
    # انجام عملیات مربوط به بارگذاری عکس

def Exit():
    window.destroy()

def show_user_info():
    first_name = First_name_varable.get()
    last_name = Last_Name_varable.get()
    major = Major_varable.get()
    date_of_birth = Date_Of_Birth_varable.get()
    skill = Skill_varable.get()
    phone_number = Phone_Number_varable.get()
    country = Country_varable.get()
    father_name = Father_Name_varable.get()


    # نمایش اطلاعات در پنجره پیغام
    message = f"First Name: {first_name}\n" \
              f"Last Name: {last_name}\n" \
              f"Major: {major}\n" \
              f"Date of Birth: {date_of_birth}\n" \
              f"Skill: {skill}\n" \
              f"Phone Number: {phone_number}\n" \
              f"Country: {country}\n" \
              f"Father Name: {father_name}"

    messagebox.showinfo("User Info", message)

def save_info_to_excel():
    first_name = First_name_varable.get()
    last_name = Last_Name_varable.get()
    major = Major_varable.get()
    date_of_birth = Date_Of_Birth_varable.get()
    skill = Skill_varable.get()
    phone_number = Phone_Number_varable.get()
    country = Country_varable.get()
    father_name = Father_Name_varable.get()

    file_path = "Student_data.xlsx"
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    row = [ID_Number_varable.get(), first_name, last_name, major, date_of_birth,
           Date_Of_Registration_varable.get(), phone_number, skill, country, father_name, "", ""]
    sheet.append(row)

    wb.save(file_path)
    messagebox.showinfo("Success", "Data saved to Excel file.")



# file = pathlib.Path("Student_data.xlsx")
#
#
# if file.exists():
#     pass
# else:
#     wb = Workbook()
#     sheet = wb.active
#     sheet["A1"] = "ID Number"
#     sheet["B1"] = "First Name"
#     sheet["C1"] = "Last Name"
#     sheet["D1"] = "Major"
#     sheet["E1"] = "Date Of Birth"
#     sheet["F1"] = "Date Of Registration"
#     sheet["G1"] = "Phone Number"
#     sheet["H1"] = "Skill"
#     sheet["I1"] = "Country"
#     sheet["J1"] = "Father Name"
#     sheet["K1"] = "Mother Name"
#     wb.save("Student_data.xlsx")

#Top_header

lable_header_creator= Label(window,
                    text="Created by Kamyar Naeimi \n Date of creation= 2023/07/22",
                    width=10,
                    height=2,
                    bg="#22343B",
                    anchor="e",
                    padx=15).pack(side=TOP,fill=X)

lable_header_main= Label(window,
                    text="STUDENT MANAGEMENT",
                    width=26,
                    height=2,
                    bg="#22343B",
                    font=("Bebas Neue",20,"bold"),
                    fg="white",
                         pady=20).pack(side=TOP,fill=X)

search= StringVar()
search_Entry= Entry(window,textvariable=search,
                    width=15,
                    bd=4,
                    font=("arial,20,bold")).place(x=950,y=100)
search_button= Button(window,
                      text="Search",
                      font="Arial 10 bold",
                      image=search_image,
                      compound=RIGHT).place(x=1100,y=100)

update_button= Button(window,
                      text="Udpate",
                      font="Arial 10 bold",
                      width=100,
                      image=update_image,
                      compound=RIGHT).place(x=180,y=70)

ID_Number_lable= Label(window,
                       text="ID Number:",
                       font="arial 13 bold",
                       bg="white").place(x=45, y=205)
Date_Of_Registration_lable= Label(window,
                       text="Date Of Registration:",
                       font="arial 13 bold",
                       bg="white").place(x=635, y=195)
#ID_Number
ID_Number_varable= StringVar()
ID_Number_entry= Entry(window,
                       textvariable=ID_Number_varable,
                       width=30,
                       font=("IMPACT 15"),
                       bd=4,
                       show="*").place(x=145, y=200)

#Date Of Registration
Date_Of_Registration_varable= StringVar()
today = date.today()
d1 = today.strftime("%d/%m/%y")
Date_Of_Registration_entry= Entry(window,
                       textvariable=Date_Of_Registration_varable,
                       width=30,
                       font=("IMPACT 15"),
                       bd=4).place(x=810, y=190)
Date_Of_Registration_varable.set(d1)

#Student Details
Student_Details_lable= Label(window,
                             font="arial 10",
                             bd=4,
                             bg="#A0AECD",
                             fg="black",
                             width=100,
                             height=25,
                             relief=RAISED).place(x=100,y=250)
obj= Label(window,text="Student's Detail-----------------------------------------------------------------------------------------------------------------------------------------------",
                            bg="#A0AECD").place(x=103,y=253)
#First name
First_name_varable=StringVar()
First_Name_lable= Label(window,
                        text="First Name:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=120, y=305)
First_Name_entry= Entry(obj,
                        textvariable=First_name_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=220, y=305)
#Last Name
Last_Name_varable=StringVar()
Last_Name_lable= Label(window,
                        text="Last Name:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=500, y=305)
Last_Name_entry= Entry(obj,
                        textvariable=Last_Name_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=600, y=305)
#Major
Major_varable=StringVar()
Major_lable= Label(window,
                        text="Major:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=120, y=360)

Major1= Combobox(obj, values=["Computer Science","Environmental Science","Electrical Engineering","Mechanical Engineering","Civil Engineering","Medicine","Business Administration","Economics","Psychology","Biology","Chemical Engineering","Other"],
                         font=("roboto,12,bold"),
                         width=20,
                         state="r").place(x=190,y=360)

Date_Of_Birth_varable=StringVar()
Date_Of_Birth_lable= Label(window,
                        text="Date Of Birth:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=500, y=360)
Date_Of_Birth_entry= Entry(obj,
                        textvariable=Date_Of_Birth_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=620, y=360)
#Skill
Skill_varable=StringVar()
Skill_lable= Label(window,
                        text="Skill:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=120, y=415)
Skill_entry= Entry(obj,
                        textvariable=Skill_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=170, y=415)
#Phone Number
Phone_Number_varable=StringVar()
Phone_Number_lable= Label(window,
                        text="Phone Number:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=500, y=415)
Phone_Number_entry= Entry(obj,
                        textvariable=Phone_Number_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=636, y=415)
#Country
Country_varable=StringVar()
Country_lable= Label(window,
                        text="Country:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=120, y=470)
Country_entry= Entry(obj,
                        textvariable=Country_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=202, y=470)
#Father Name
Father_Name_varable=StringVar()
Father_Name_lable= Label(window,
                        text="Father Name:",
                        bd=5,
                        bg="#A0AECD",
                        font="arial 13 bold").place(x=500, y=470)
Father_Name_entry= Entry(obj,
                        textvariable=Father_Name_varable,
                        width=20,
                        relief=RAISED,
                        font="arial 15 bold",
                        background="white").place(x=620, y=470)
#
# gender= IntVar()
# gender1= Radiobutton(obj,
#                      text="Male",
#                      value=1,
#                      command=select_gender,
#                      variable=gender,
#                      bg="white").place(x=120, y=490)
# gender2= Radiobutton(obj,
#                      text="Female",
#                      value=2,
#                      command=select_gender,
#                      variable=gender,
#                      bg="white").place(x=170, y=490)
# gender3= Radiobutton(obj,
#                      text="Other",
#                      value=3,
#                      command=select_gender,
#                      variable=gender,
#                      bg="white").place(x=220, y=490)
# gender4= Radiobutton(obj,
#                      text="Prefer Not To Say",
#                      value=4,
#                      command=select_gender,
#                      variable=gender,
#                      bg="white").place(x=270, y=490)

Information_text= Label(window,text="Show Details-----------------------------------------------------------------------------------------------------------------------------------------------",
                            bg="#A0AECD").place(x=103,y=510)

profile_icon= PhotoImage(file="large_profile_icon.png")
f= Frame(window,
         bg="black",
         width=200,height=200,
         relief=GROOVE,
         bd=5).place(x=970,y=250)

profile_lable= Label(f,bg="black",image=profile_icon,background="#7DA2A9",
                     relief=FLAT).place(x=970,y=250)

Upload_button= Button(window, text="Upload",
                      font="arial 15 bold",
                      height=1,
                      width=10,
                      bg="#161F6D",
                      command=Upload_photo).place(x=1000, y=460)

Reset_button= Button(window, text="Reset",
                      font="arial 15 bold",
                      height=1,
                      width=10,
                      bg="#161F6D").place(x=1000, y=510)

save_button = Button(window,  bg="#161F6D",text="Save to Excel", font="Arial 15 bold", command=save_info_to_excel)
save_button.place(x=1000, y=510)

Exit_button= Button(window, text="Exit",
                      font="arial 15 bold",
                      height=1,
                      width=10,
                      bg="#161F6D",
                    command=Exit).place(x=1000, y=610)

show_info_button = Button(window,  bg="#161F6D",text="Show Info", font="Arial 15 bold", command=show_user_info)
show_info_button.place(x=1000, y=560)

window.mainloop()