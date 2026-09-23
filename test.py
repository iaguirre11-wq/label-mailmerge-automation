import win32com.client

template_path = r"C:\Users\AguirreIan\OneDrive - Suffern Central School District\Documents\P-Touch\Labels\Student Chromebook Label.lbx"
preview_path = r"C:\Users\AguirreIan\Documents\Suffern Odd projects\Label Automation\test_output.bmp"


# Fake test values — stand-ins for your six form fields
student_name = "Test Student that is very long"
school = "Montebello Elementary"
yog = "2036"
serial_number = "PF6C49RA"

doc = win32com.client.Dispatch("bpac.Document")

if not doc.Open(template_path):
    print("Failed to open template.")
    raise SystemExit

try:

    doc.GetObject("Name").Text = student_name
    doc.GetObject("School").Text = school
    doc.GetObject("YOG").Text = yog
    doc.GetObject("Bar Code").Text = serial_number
    doc.GetObject("QR Code").Text = serial_number

 # --- NEW: choose printer instead of using the one saved in the .lbx ---
    print("Template's saved printer:", doc.Printer.Name)

    installed = doc.Printer.GetInstalledPrinters
    brother = [p for p in installed if doc.Printer.IsPrinterSupported(p)]

    for i, p in enumerate(brother):
        print(i, p)

    choice = int(input("Pick printer number: "))
    print("SetPrinter:", doc.SetPrinter(brother[choice], False))

    # Which printer will b-PAC send to? (the one saved in the .lbx)
    printer_name = doc.Printer.Name
    print("Target printer:", printer_name)
    print("Online:", doc.Printer.IsPrinterOnline(printer_name))

    # Preview first, same as before
    print("Preview export:", doc.Export(4, preview_path, 180))

    if input("Print one real label? (y/n): ").strip().lower() == "y":
        started = doc.StartPrint("", 0)
        print("StartPrint:", started)

        if started:
            printed = doc.PrintOut(1, 0)
            print("PrintOut:", printed)
            ended = doc.EndPrint          # no parentheses
            print("EndPrint:", ended)

        if not started or not printed:
            print("Error code:", doc.ErrorCode)
    else:
        print("Skipped printing.")

finally:
    doc.Close
