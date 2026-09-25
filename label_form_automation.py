import openpyxl
import win32com.client
import win32print
import sys
import tomllib
import pywintypes
from pathlib import Path
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

# b-PAC is used with late binding (win32com.client.Dispatch).
# gencache.EnsureDispatch does NOT work with b-PAC (makepy error), so:
#   - Methods that take NO arguments are invoked automatically when accessed,
#     so they are written WITHOUT parentheses (doc.Close, doc.EndPrint,
#     doc.Printer.GetInstalledPrinters).
#     Adding () causes "TypeError: 'bool'/'tuple' object is not callable".
#   - Methods WITH arguments keep their parentheses (doc.Open(path), doc.PrintOut(1, 0)).
#   - Properties never use parentheses (doc.Printer.Name).


def fatal_startup_error(title, message):
    temp_root = Tk()
    temp_root.withdraw()
    messagebox.showerror(title, message)
    temp_root.destroy()
    sys.exit(1)


if getattr(sys, "frozen", False):
    base_dir = Path(sys.executable).parent
else:
    base_dir = Path(__file__).parent

config_path = base_dir / "config.toml"

try:
    with open(config_path, "rb") as f:
        config = tomllib.load(f)
    # Configuration paths and settings loaded from config.toml
    excel_database = config["paths"]["excel_database"]
    label_template = config["paths"]["label_template"]
    form_template = config["paths"]["form_template"]
    school_codes = config["schools"]
    device_models = config["devices"]["models"]
except FileNotFoundError:
    fatal_startup_error("Config Missing",
                        f"Couldn't find the config file:\n{config_path}\n\n"
                        "Copy config.example.toml to config.toml in the same folder and fill in your paths.")

except tomllib.TOMLDecodeError as e:
    fatal_startup_error("Config Error",
                        f"config.toml has a formatting error:\n{e}\n\n"
                        "Tip: Windows paths must be in single quotes.")

except KeyError as e:
    fatal_startup_error("Config Error",
                        f"config.toml is missing a setting: {e}\n\n"
                        "Compare it against config.example.toml.")

# Placeholders for combobox selection
no_printer_placeholder = "No Printers Available"
school_code_placeholder = "Select a school"
device_model_placeholder = "Select a device model"


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
    installed = doc.Printer.GetInstalledPrinters or []  # no ()
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

        if not doc.SetPrinter(printer_name, False):
            raise RuntimeError(
                f"Failed to set label printer to '{printer_name}' (b-PAC error {doc.ErrorCode})")

        if not doc.StartPrint("", 0):
            raise RuntimeError(
                f"Failed to start printing label (b-PAC error {doc.ErrorCode})")

        if not doc.PrintOut(1, 0):
            raise RuntimeError(
                f"Failed to print label (b-PAC error {doc.ErrorCode})")

        if not doc.EndPrint:  # no ()
            raise RuntimeError(
                f"Failed to end printing label (b-PAC error {doc.ErrorCode})")

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

    problems = record_problems(record)

    if problems:
        messagebox.showwarning("Can't Submit Info!", "\n".join(problems))
        return

    if not do_label.get() and not do_form.get() and not do_database.get():
        messagebox.showwarning(
            "No Actions Selected", "Please select at least one action to perform.")
        return

    if do_label.get():
        if combo_printer.get() == no_printer_placeholder:
            messagebox.showwarning(
                "Printer Warning", "No Brother label printers found. Check that the printer is installed and b-PAC is set up, then restart the tool.")
            return

    if do_form.get():
        if combo_regular_printer.get() == no_printer_placeholder:
            messagebox.showwarning(
                "Printer Warning", "No form/regular printers available.")
            return

    if do_label.get():
        try:
            print_label(record, combo_printer.get())
        except Exception as e:
            messagebox.showerror("Label Error", str(e))
            return

    if do_form.get():
        try:
            print_form(record, combo_regular_printer.get())
        except Exception as e:
            messagebox.showerror("Form Error", str(e))
            return

    if do_database.get():
        try:
            write_to_excel(record)
        except Exception as e:
            messagebox.showerror(
                "Excel Error", f"The spreadsheet is open or locked. Close it, uncheck the other actions, and submit again to save just the row.\n" + str(e))
            return

    do_label.set(True)
    do_form.set(True)
    do_database.set(True)


def record_problems(record):

    warning_list = []

    if record["selected_school"] not in school_codes:
        warning_list.append("Please select a school.")

    if record["selected_model"] not in device_models:
        warning_list.append("Please select a device model.")

    return warning_list


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
combo_school.set(school_code_placeholder)
combo_school.grid(column=2, row=5, sticky=(W, E))

ttk.Label(mainframe, text="School:").grid(column=1, row=5, sticky=W)

# Section for Device Model
model_list = list(device_models)

combo_device = ttk.Combobox(
    mainframe, width=30, values=model_list, state="readonly")
combo_device.set(device_model_placeholder)
combo_device.grid(column=2, row=6, sticky=(W, E))

ttk.Label(mainframe, text="Device Model:").grid(column=1, row=6, sticky=W)

# Populating the printer list
try:
    printer_list = get_brother_printers()
except pywintypes.com_error:
    messagebox.showwarning(
        "Label Printer Warning",
        "The Brother b-PAC SDK could not be loaded, so labels can't be printed.\n\n"
        "Install the 64-bit b-PAC SDK, then restart the tool.")
    printer_list = []

# Section for Selecting Brother Printer
combo_printer = ttk.Combobox(mainframe, values=printer_list, state="readonly")
if printer_list:
    combo_printer.set(printer_list[0])

else:
    combo_printer.set(no_printer_placeholder)

combo_printer.grid(column=2, row=7, sticky=(W, E))

ttk.Label(mainframe, text="Label Printer:").grid(column=1, row=7, sticky=W)

# Section for Selecting Regular Printer
regular_printer_list = get_form_printers(printer_list)

combo_regular_printer = ttk.Combobox(
    mainframe, values=regular_printer_list, state="readonly")

try:
    default_printer = win32print.GetDefaultPrinter()
except pywintypes.error:
    default_printer = None

if default_printer in regular_printer_list:
    combo_regular_printer.set(default_printer)
elif regular_printer_list:
    combo_regular_printer.set(regular_printer_list[0])
else:
    combo_regular_printer.set(no_printer_placeholder)

combo_regular_printer.grid(column=2, row=8, sticky=(W, E))
ttk.Label(mainframe, text="Form Printer:").grid(column=1, row=8, sticky=W)

# Section for selecting print options
do_label = BooleanVar(value=True)
ttk.Checkbutton(mainframe, text="Print Labels",
                variable=do_label).grid(column=3, row=6, sticky=W)

do_form = BooleanVar(value=True)
ttk.Checkbutton(mainframe, text="Print Forms",
                variable=do_form).grid(column=3, row=7, sticky=W)

do_database = BooleanVar(value=True)
ttk.Checkbutton(mainframe, text="Save to Database",
                variable=do_database).grid(column=3, row=8, sticky=W)

# Section for Submit Button
submit_button = ttk.Button(mainframe, text="Submit", command=submit_info)
submit_button.grid(column=2, row=10, sticky=W)

root.mainloop()
# -------------------------------------------------------------------------
