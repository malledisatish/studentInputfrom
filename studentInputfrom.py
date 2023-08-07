import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

def enter_data():
    accepted=accept_var.get()

    if accepted=="Accepted":
        student_name=student_name_entry.get()
        parent_name=parent_name_entry.get()
        if student_name and parent_name:
            course=courses_combobox.get()
            age=age_spinbox.get()
            nationality=nationality_combobox.get()
            college=college_spinbox.get()
            NumCourses=courses_spinbox.get()
            registered_status=registered_var.get()

            print("Student name:",student_name,"parent name:", parent_name)
            print("courses: ",course,"Age:",age,"Nationality:",nationality)
            print("college:",college,"No.of courses",NumCourses)
            print("Registered status:", registered_status)

            filepath="E:\INTERVIEW\data.xlsx"
            if not os.path.exists(filepath):
                workbook=openpyxl.Workbook()
                sheet=workbook.active
                heading=["Student Name","Parent Name","Course","Age","Nationality","College","No.of courses","Registered status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook=openpyxl.load_workbook(filepath)
            sheet=workbook.active
            sheet.append([student_name,parent_name,course,age,nationality,college,NumCourses,registered_status])
            workbook.save(filepath)


        else:
            tkinter.messagebox.showwarning(title="Error", message="Enter First name and last name")
    else:
        tkinter.messagebox.showwarning(title="Error",message="You have not accepted")

window=tkinter.Tk()
window.title("Student Input Form")

frame=tkinter.Frame(window)
frame.pack()

student_info_frame=tkinter.LabelFrame(frame,text="Student Information")
student_info_frame.grid(row=0,column=0)

student_name_label=tkinter.Label(student_info_frame,text="Student Name")
student_name_label.grid(row=0,column=0)

parent_name_label=tkinter.Label(student_info_frame,text="Parent Name")
parent_name_label.grid(row=0,column=1)

student_name_entry=tkinter.Entry(student_info_frame)
parent_name_entry=tkinter.Entry(student_info_frame)
student_name_entry.grid(row=1,column=0)
parent_name_entry.grid(row=1,column=1)

courses_label=tkinter.Label(student_info_frame,text="Courses")
courses_combobox=ttk.Combobox(student_info_frame,values=["English","Telugu","Hindi","Maths","Science","Social"])
courses_label.grid(row=0,column=2)
courses_combobox.grid(row=1,column=2)

age_label=tkinter.Label(student_info_frame,text="Age")
age_spinbox=tkinter.Spinbox(student_info_frame, from_=18,to=110)
age_label.grid(row=2,column=0)
age_spinbox.grid(row=3,column=0)

nationality_label=tkinter.Label(student_info_frame,text="Nationality")
nationality_combobox=ttk.Combobox(student_info_frame,values=["India","USA","UK"])
nationality_label.grid(row=2,column=1)
nationality_combobox.grid(row=3,column=1)

for widget in student_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

courses_frame=tkinter.LabelFrame(frame)
courses_frame.grid(row=1,column=0,sticky="news",padx=20,pady=10)

registered_label=tkinter.Label(courses_frame,text="Registration")
registered_var=tkinter.StringVar(value="Not registered")
registered_check=tkinter.Checkbutton(courses_frame,text="Currently Registered",
                                     variable=registered_var, onvalue="Registered", offvalue="Not registered")
registered_label.grid(row=0,column=0)
registered_check.grid(row=1,column=0)

college_label=tkinter.Label(courses_frame,text="College Code")
college_spinbox=tkinter.Spinbox(courses_frame,from_=0,to="infinity")
college_label.grid(row=0,column=1)
college_spinbox.grid(row=1,column=1)

courses_label=tkinter.Label(courses_frame,text="No.of Courses")
courses_spinbox=tkinter.Spinbox(courses_frame,from_=0,to="infinity")
courses_label.grid(row=0,column=2)
courses_spinbox.grid(row=1,column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

terms_frame=tkinter.LabelFrame(frame,text="Terms & Conditions")
terms_frame.grid(rows=2,column=0,sticky="news",padx=20, pady=10)

accept_var=tkinter.StringVar(value="")
term_check=tkinter.Checkbutton(terms_frame,text="I accept the terms and conditions.",variable=accept_var,onvalue="Accepted",offvalue="Not Accepted")
term_check.grid(row=0,column=0)

button=tkinter.Button(frame,text="Enter data",command=enter_data)
button.grid(row=4,column=0,sticky="news",padx=20,pady=10)



window.mainloop()
