import tkinter as tk
from tkinter import messagebox
import openpyxl
import os
from datetime import date

FILE_NAME = "attendance_records.xlsx"

COLUMNS = [
    "Record ID",
    "Student ID",
    "Student Name",
    "Class",
    "Date",
    "Subject",
    "Attendance"
]


# ---------------- CREATE EXCEL FILE ----------------

def create_excel_file():

    if not os.path.exists(FILE_NAME):

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"

        ws.append(COLUMNS)

        wb.save(FILE_NAME)
        wb.close()


# ---------------- CLEAR FORM ----------------

def clear_form():

    student_id_entry.delete(0, tk.END)
    name_entry.delete(0, tk.END)
    class_entry.delete(0, tk.END)
    subject_entry.delete(0, tk.END)
    search_entry.delete(0, tk.END)

    date_entry.delete(0, tk.END)
    date_entry.insert(0, str(date.today()))

    attendance_var.set("Present")


# ---------------- GET NEXT ID ----------------

def get_next_id():

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    last_id = 0

    for row in ws.iter_rows(min_row=2, values_only=True):

        if row[0] is not None:

            try:
                last_id = max(last_id, int(row[0]))
            except:
                pass

    wb.close()

    return last_id + 1


# ---------------- ADD ATTENDANCE ----------------

def add_record():

    student_id = student_id_entry.get().strip()
    student_name = name_entry.get().strip()
    student_class = class_entry.get().strip()
    attendance_date = date_entry.get().strip()
    subject = subject_entry.get().strip()
    attendance = attendance_var.get()

    if (student_id == "" or
        student_name == "" or
        student_class == "" or
        attendance_date == "" or
        subject == ""):

        messagebox.showwarning(
            "Missing Data",
            "Please fill all fields."
        )

        return

    record_id = get_next_id()

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    ws.append([
        record_id,
        student_id,
        student_name,
        student_class,
        attendance_date,
        subject,
        attendance
    ])

    wb.save(FILE_NAME)
    wb.close()

    messagebox.showinfo(
        "Success",
        "Attendance added successfully!"
    )

    clear_form()


# ---------------- SHOW ALL RECORDS ----------------

def show_records():

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    records = ""

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        records += (
            "Record ID: " + str(row[0]) +
            "\nStudent ID: " + str(row[1]) +
            "\nStudent Name: " + str(row[2]) +
            "\nClass: " + str(row[3]) +
            "\nDate: " + str(row[4]) +
            "\nSubject: " + str(row[5]) +
            "\nAttendance: " + str(row[6]) +
            "\n--------------------------\n"
        )

    wb.close()

    if records == "":
        records = "No attendance records found."

    messagebox.showinfo(
        "All Attendance Records",
        records
    )


# ---------------- SEARCH ----------------

def search_record():

    search_value = search_entry.get().strip().lower()

    if search_value == "":

        messagebox.showwarning(
            "Search",
            "Enter Student ID or Student Name."
        )

        return

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    records = ""

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        student_id = str(row[1]).lower()
        student_name = str(row[2]).lower()

        if (search_value in student_id or
            search_value in student_name):

            records += (
                "Record ID: " + str(row[0]) +
                "\nStudent ID: " + str(row[1]) +
                "\nStudent Name: " + str(row[2]) +
                "\nClass: " + str(row[3]) +
                "\nDate: " + str(row[4]) +
                "\nSubject: " + str(row[5]) +
                "\nAttendance: " + str(row[6]) +
                "\n--------------------------\n"
            )

    wb.close()

    if records == "":
        records = "No matching record found."

    messagebox.showinfo(
        "Search Result",
        records
    )


# ---------------- FIND RECORD FOR UPDATE/DELETE ----------------

def find_record():

    student_id = student_id_entry.get().strip()

    if student_id == "":

        messagebox.showwarning(
            "Find Record",
            "Enter Student ID first."
        )

        return

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    found = False

    for row in ws.iter_rows(min_row=2):

        if str(row[1].value) == student_id:

            name_entry.delete(0, tk.END)
            name_entry.insert(0, str(row[2].value))

            class_entry.delete(0, tk.END)
            class_entry.insert(0, str(row[3].value))

            date_entry.delete(0, tk.END)
            date_entry.insert(0, str(row[4].value))

            subject_entry.delete(0, tk.END)
            subject_entry.insert(0, str(row[5].value))

            attendance_var.set(str(row[6].value))

            found = True
            break

    wb.close()

    if not found:

        messagebox.showinfo(
            "Find Record",
            "Student record not found."
        )


# ---------------- UPDATE ----------------

def update_record():

    student_id = student_id_entry.get().strip()
    student_name = name_entry.get().strip()
    student_class = class_entry.get().strip()
    attendance_date = date_entry.get().strip()
    subject = subject_entry.get().strip()
    attendance = attendance_var.get()

    if student_id == "":

        messagebox.showwarning(
            "Update",
            "Enter Student ID first."
        )

        return

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    found = False

    for row in ws.iter_rows(min_row=2):

        if str(row[1].value) == student_id:

            row[2].value = student_name
            row[3].value = student_class
            row[4].value = attendance_date
            row[5].value = subject
            row[6].value = attendance

            found = True
            break

    if found:

        wb.save(FILE_NAME)
        wb.close()

        messagebox.showinfo(
            "Success",
            "Attendance record updated successfully!"
        )

        clear_form()

    else:

        wb.close()

        messagebox.showinfo(
            "Update",
            "Student record not found."
        )


# ---------------- DELETE ----------------

def delete_record():

    student_id = student_id_entry.get().strip()

    if student_id == "":

        messagebox.showwarning(
            "Delete",
            "Enter Student ID first."
        )

        return

    confirm = messagebox.askyesno(
        "Confirm Delete",
        "Are you sure you want to delete this record?"
    )

    if not confirm:
        return

    wb = openpyxl.load_workbook(FILE_NAME)
    ws = wb.active

    found = False

    for row_number in range(
        2,
        ws.max_row + 1
    ):

        if str(ws.cell(
            row_number,
            2
        ).value) == student_id:

            ws.delete_rows(
                row_number,
                1
            )

            found = True
            break

    wb.save(FILE_NAME)
    wb.close()

    if found:

        messagebox.showinfo(
            "Deleted",
            "Attendance record deleted successfully!"
        )

        clear_form()

    else:

        messagebox.showinfo(
            "Delete",
            "Student record not found."
        )


# ---------------- MAIN WINDOW ----------------

create_excel_file()

root = tk.Tk()

root.title(
    "Students Attendance Management System"
)

root.geometry("400x750")

root.configure(
    bg="#EAF4F4"
)


# ---------------- TITLE ----------------

title = tk.Label(
    root,
    text="Students Attendance\nManagement System",
    font=("Arial", 17, "bold"),
    bg="#145A32",
    fg="white",
    pady=15
)

title.pack(fill="x")


# ---------------- FORM ----------------

form = tk.Frame(
    root,
    bg="#EAF4F4"
)

form.pack(
    fill="both",
    expand=True,
    padx=20,
    pady=12
)


# Student ID

tk.Label(
    form,
    text="Student ID",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")

student_id_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

student_id_entry.pack(
    fill="x",
    pady=(2, 6)
)


# Student Name

tk.Label(
    form,
    text="Student Name",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")

name_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

name_entry.pack(
    fill="x",
    pady=(2, 6)
)


# Class

tk.Label(
    form,
    text="Class",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")

class_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

class_entry.pack(
    fill="x",
    pady=(2, 6)
)


# Date

tk.Label(
    form,
    text="Date",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")

date_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

date_entry.pack(
    fill="x",
    pady=(2, 6)
)

date_entry.insert(
    0,
    str(date.today())
)


# Subject

tk.Label(
    form,
    text="Subject",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")

subject_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

subject_entry.pack(
    fill="x",
    pady=(2, 6)
)


# Attendance

tk.Label(
    form,
    text="Attendance",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(anchor="w")


attendance_var = tk.StringVar()

attendance_var.set("Present")


attendance_menu = tk.OptionMenu(
    form,
    attendance_var,
    "Present",
    "Absent"
)

attendance_menu.config(
    font=("Arial", 12)
)

attendance_menu.pack(
    fill="x",
    pady=(2, 8)
)


# ---------------- BUTTONS ----------------

tk.Button(
    form,
    text="ADD ATTENDANCE",
    font=("Arial", 10, "bold"),
    bg="#27AE60",
    fg="white",
    command=add_record
).pack(
    fill="x",
    pady=2
)


tk.Button(
    form,
    text="FIND RECORD",
    font=("Arial", 10, "bold"),
    bg="#3498DB",
    fg="white",
    command=find_record
).pack(
    fill="x",
    pady=2
)


tk.Button(
    form,
    text="UPDATE RECORD",
    font=("Arial", 10, "bold"),
    bg="#F39C12",
    fg="white",
    command=update_record
).pack(
    fill="x",
    pady=2
)


tk.Button(
    form,
    text="DELETE RECORD",
    font=("Arial", 10, "bold"),
    bg="#E74C3C",
    fg="white",
    command=delete_record
).pack(
    fill="x",
    pady=2
)


# ---------------- SEARCH ----------------

tk.Label(
    form,
    text="Search Student",
    font=("Arial", 12, "bold"),
    bg="#EAF4F4"
).pack(
    anchor="w",
    pady=(7, 0)
)


search_entry = tk.Entry(
    form,
    font=("Arial", 12)
)

search_entry.pack(
    fill="x",
    pady=2
)


tk.Button(
    form,
    text="SEARCH",
    font=("Arial", 10, "bold"),
    bg="#8E44AD",
    fg="white",
    command=search_record
).pack(
    fill="x",
    pady=2
)


# ---------------- SHOW ALL ----------------

tk.Button(
    form,
    text="SHOW ALL RECORDS",
    font=("Arial", 10, "bold"),
    bg="#2980B9",
    fg="white",
    command=show_records
).pack(
    fill="x",
    pady=2
)


# ---------------- CLEAR ----------------

tk.Button(
    form,
    text="CLEAR",
    font=("Arial", 10, "bold"),
    command=clear_form
).pack(
    fill="x",
    pady=2
)


# ---------------- RUN ----------------

root.mainloop()