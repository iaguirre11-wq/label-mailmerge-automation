from tkinter import *
from tkinter import ttk


def submit_info():
    student_name = name.get()
    year_of_graduation = YOG.get()
    student_grade = grade.get()
    device_serial_number = serial_number.get()
    selected_school = combo.get()
    selected_model = combo.get()

    print("Student Name:", student_name)
    print("Year of Graduation:", year_of_graduation)
    print("Grade:", student_grade)
    print("Device Serial Number:", device_serial_number)
    print("School:", selected_school)
    print("Device Model:", selected_model)


root = Tk()
root.title("Student Information Form")

mainframe = ttk.Frame(root, padding=(3, 3, 12, 12))
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))

# Section for entering student name:
name = StringVar()
name_entry = ttk.Entry(mainframe, width=20, textvariable=name)
name_entry.grid(column=2, row=1, sticky=(W, E))

ttk.Label(mainframe, text="Student Name:").grid(column=1, row=1, sticky=W)

# Section for entering year of graduation:
YOG = StringVar()
YOG_entry = ttk.Entry(mainframe, width=20, textvariable=YOG)
YOG_entry.grid(column=2, row=2, sticky=(W, E))

ttk.Label(mainframe, text="Year of Graduation:").grid(
    column=1, row=2, sticky=W)

# Section for entering grade:
grade = StringVar()
grade_entry = ttk.Entry(mainframe, width=20, textvariable=grade)
grade_entry.grid(column=2, row=3, sticky=(W, E))

ttk.Label(mainframe, text="Grade:").grid(column=1, row=3, sticky=W)

# Section for New Device Serial Number:
serial_number = StringVar()
serial_entry = ttk.Entry(mainframe, width=20, textvariable=serial_number)
serial_entry.grid(column=2, row=4, sticky=(W, E))

ttk.Label(mainframe, text="Device Serial Number:").grid(
    column=1, row=4, sticky=W)

# Section for entering School
school_list = ["School A", "School B", "School C"]

combo = ttk.Combobox(mainframe, values=school_list, state="readonly")
combo.set("Select School")
combo.grid(column=2, row=5, sticky=(W, E))

ttk.Label(mainframe, text="School:").grid(column=1, row=5, sticky=W)

# Section for Device Model
model_list = ["Model X", "Model Y", "Model Z"]

combo = ttk.Combobox(mainframe, values=model_list, state="readonly")
combo.set("Select Model")
combo.grid(column=2, row=6, sticky=(W, E))

ttk.Label(mainframe, text="Device Model:").grid(column=1, row=6, sticky=W)

submit_button = ttk.Button(mainframe, text="Submit", command=submit_info)
submit_button.grid(column=2, row=7, sticky=W)


root.mainloop()
