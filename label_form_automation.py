import openpyxl
import win32com.client
from tkinter import *
from tkinter import ttk


def write_to_excel(record):
    workbook = openpyxl.load_workbook(
        r'C:\Users\AguirreIan\OneDrive - Suffern Central School District\Documents\P-Touch\Databases\Test1.xlsx')

    worksheet = workbook.active
    last_row = worksheet.max_row + 1

    worksheet.cell(row=last_row, column=1,
                   value=record["device_serial_number"])
    worksheet.cell(row=last_row, column=2, value=record["student_name"])
    worksheet.cell(row=last_row, column=3, value=record["year_of_graduation"])
    worksheet.cell(row=last_row, column=4, value=record["student_grade"])
    worksheet.cell(row=last_row, column=5, value=record["selected_school"])
    worksheet.cell(row=last_row, column=6, value=record["selected_model"])

    workbook.save(
        r'C:\Users\AguirreIan\OneDrive - Suffern Central School District\Documents\P-Touch\Databases\Test1.xlsx')


def submit_info():

    record = {
        "device_serial_number": serial_number.get(),
        "student_name": name.get(),
        "year_of_graduation": YOG.get(),
        "student_grade": grade.get(),
        "selected_school": combo_school.get(),
        "selected_model": combo_device.get()
    }

    write_to_excel(record)


# --------User GUI--------------------
root = Tk()
root.title("Student Information Form")

mainframe = ttk.Frame(root, padding=(3, 3, 12, 12))
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))

# Section for New Device Serial Number:
serial_number = StringVar()
serial_entry = ttk.Entry(mainframe, width=20, textvariable=serial_number)
serial_entry.grid(column=2, row=1, sticky=(W, E))

ttk.Label(mainframe, text="Device Serial Number:").grid(
    column=1, row=1, sticky=W)

# Section for entering student name:
name = StringVar()
name_entry = ttk.Entry(mainframe, width=20, textvariable=name)
name_entry.grid(column=2, row=2, sticky=(W, E))

ttk.Label(mainframe, text="Student Name:").grid(column=1, row=2, sticky=W)

# Section for entering year of graduation:
YOG = StringVar()
YOG_entry = ttk.Entry(mainframe, width=20, textvariable=YOG)
YOG_entry.grid(column=2, row=3, sticky=(W, E))

ttk.Label(mainframe, text="Year of Graduation:").grid(
    column=1, row=3, sticky=W)

# Section for entering grade:
grade = StringVar()
grade_entry = ttk.Entry(mainframe, width=20, textvariable=grade)
grade_entry.grid(column=2, row=4, sticky=(W, E))

ttk.Label(mainframe, text="Grade:").grid(column=1, row=4, sticky=W)

# Section for entering School
school_list = ["School A", "School B", "School C"]

combo_school = ttk.Combobox(mainframe, values=school_list, state="readonly")
combo_school.set("Select School")
combo_school.grid(column=2, row=5, sticky=(W, E))

ttk.Label(mainframe, text="School:").grid(column=1, row=5, sticky=W)

# Section for Device Model
model_list = ["Model X", "Model Y", "Model Z"]

combo_device = ttk.Combobox(mainframe, values=model_list, state="readonly")
combo_device.set("Select Model")
combo_device.grid(column=2, row=6, sticky=(W, E))

ttk.Label(mainframe, text="Device Model:").grid(column=1, row=6, sticky=W)

submit_button = ttk.Button(mainframe, text="Submit", command=submit_info)
submit_button.grid(column=2, row=7, sticky=W)


root.mainloop()
# -------------------------------------------------------------------------
