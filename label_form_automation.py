import openpyxl
import win32com.client
import win32print
import sys
import tomllib
from pathlib import Path
from tkinter import *
from tkinter import ttk

# b-PAC is used with late binding (win32com.client.Dispatch).
# gencache.EnsureDispatch does NOT work with b-PAC (makepy error), so:
#   - Methods that take NO arguments are invoked automatically when accessed,
#     so they are written WITHOUT parentheses (doc.Close, doc.EndPrint,
#     doc.Printer.GetInstalledPrinters).
#     Adding () causes "TypeError: 'bool'/'tuple' object is not callable".
#   - Methods WITH arguments keep their parentheses (doc.Open(path), doc.PrintOut(1, 0)).
#   - Properties never use parentheses (doc.Printer.Name).

if getattr(sys, "frozen", False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent

config_path = base_dir / "config.toml"
with open(config_path, "rb") as f:
    config = tomllib.load(f)

# temp hardcoded values
excel_database = config["paths"]["excel_database"]
label_template = config["paths"]["label_template"]
form_template = config["paths"]["form_template"]
school_codes = config["schools"]
device_models = config["devices"]["models"]


def write_to_excel(record):
    workbook = openpyxl.load_workbook(excel_database)
    worksheet = workbook.active

# Append a new row to the Excel worksheet with the record data in the same order as listed below starting from column A
    worksheet.append([
        record["device_serial_number"],
        record["student_name"],
        record["year_of_graduation"],
        record["student_grade"],
        record["selected_school"],
        record["selected_model"]
    ])

    workbook.save(excel_database)


def get_brother_printers():
    doc = win32com.client.Dispatch("bpac.Document")
    installed = doc.Printer.GetInstalledPrinters  # no ()
    return [p for p in installed if doc.Printer.IsPrinterSupported(p)]


def get_form_printers(brother_printers):
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    all_printers = [p[2] for p in win32print.EnumPrinters(flags)]
    brother = set(brother_printers)
    return [p for p in all_printers if p not in brother]


def print_label(record, printer_name):
    doc = win32com.client.Dispatch("bpac.Document")

    if not doc.Open(label_template):
        raise RuntimeError(f"Could not open label template: {label_template}")

    try:

        doc.GetObject("Name").Text = record["student_name"]
        doc.GetObject("School").Text = record["selected_school"]
        doc.GetObject("YOG").Text = record["year_of_graduation"]
        doc.GetObject("Bar Code").Text = record["device_serial_number"]
        doc.GetObject("QR Code").Text = record["device_serial_number"]

        doc.SetPrinter(printer_name, False)

        doc.StartPrint("", 0)
        doc.PrintOut(1, 0)
        doc.EndPrint  # no ()

    finally:
        doc.Close  # no ()


def print_form(record, printer_name):

    field_values = {
        "User": record["student_name"],
        "Grade": record["student_grade"],
        "Device": record["selected_model"],
        "Serial": record["device_serial_number"],
        "SchoolCode": school_codes.get(record["selected_school"], ""),
    }

    word = None
    doc = None
    original_printer = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        original_printer = word.ActivePrinter
        word.ActivePrinter = printer_name
        doc = word.Documents.Open(form_template, ReadOnly=True)

        for story in doc.StoryRanges:
            while story is not None:
                for i in range(story.Fields.Count, 0, -1):
                    field = story.Fields(i)
                    parts = field.Code.Text.split()
                    if len(parts) >= 2 and parts[0] == "MERGEFIELD" and parts[1] in field_values:
                        field.Result.Text = field_values[parts[1]]
                        field.Unlink()
                story = story.NextStoryRange

        doc.PrintOut(Background=False)

    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        if word is not None:
            if original_printer:
                word.ActivePrinter = original_printer
            word.Quit()


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

    # print_label(record, combo_printer.get())
    print_form(record, combo_regular_printer.get())


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
school_list = list(school_codes)

combo_school = ttk.Combobox(mainframe, values=school_list, state="readonly")
combo_school.set("Select School")
combo_school.grid(column=2, row=5, sticky=(W, E))

ttk.Label(mainframe, text="School:").grid(column=1, row=5, sticky=W)

# Section for Device Model
model_list = list(device_models)

combo_device = ttk.Combobox(mainframe, values=model_list, state="readonly")
combo_device.set("Select Model")
combo_device.grid(column=2, row=6, sticky=(W, E))

ttk.Label(mainframe, text="Device Model:").grid(column=1, row=6, sticky=W)

# Section for Selecting Brother Printer
printer_list = get_brother_printers()

combo_printer = ttk.Combobox(mainframe, values=printer_list, state="readonly")
if printer_list:
    combo_printer.set(printer_list[0])

else:
    combo_printer.set("No Printers Available")

combo_printer.grid(column=2, row=7, sticky=(W, E))

ttk.Label(mainframe, text="Label Printer:").grid(column=1, row=7, sticky=W)

# Section for Selecting Regular Printer
regular_printer_list = get_form_printers(printer_list)

combo_regular_printer = ttk.Combobox(
    mainframe, values=regular_printer_list, state="readonly")

default_printer = win32print.GetDefaultPrinter()

if default_printer in regular_printer_list:
    combo_regular_printer.set(default_printer)
elif regular_printer_list:
    combo_regular_printer.set(regular_printer_list[0])
else:
    combo_regular_printer.set("No Printers Available")

combo_regular_printer.grid(column=2, row=8, sticky=(W, E))
ttk.Label(mainframe, text="Form Printer:").grid(column=1, row=8, sticky=W)

# Section for Submit Button
submit_button = ttk.Button(mainframe, text="Submit", command=submit_info)
submit_button.grid(column=2, row=9, sticky=W)

root.mainloop()
# -------------------------------------------------------------------------
