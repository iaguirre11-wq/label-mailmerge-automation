import win32com.client

template_path = r"C:\Users\AguirreIan\Documents\Suffern Odd projects\Label Automation\2026-2027_Student Device Agreement.docx"   # your prepped copy
pdf_preview = r"C:\Users\AguirreIan\Documents\Suffern Odd projects\Label Automation\form_preview.pdf"

SCHOOL_CODES = {
    "Suffern High School": "SHS",
    "Montebello Elementary": "MES",
    # fill in the rest
}

record = {
    "device_serial_number": "PF6C49RA",
    "student_name": "Test Student",
    "student_grade": "5",
    "selected_school": "Montebello Elementary",
    "selected_model": "Test Model",
}

values = {
    "User": record["student_name"],
    "Grade": record["student_grade"],
    "Device": record["selected_model"],
    "Serial": record["device_serial_number"],
    "SchoolCode": SCHOOL_CODES.get(record["selected_school"], ""),
}

word = win32com.client.DispatchEx("Word.Application")   # separate hidden Word
word.Visible = False
word.DisplayAlerts = 0

doc = word.Documents.Open(template_path, ReadOnly=True, AddToRecentFiles=False)
try:
    for story in doc.StoryRanges:              # body, headers, footers...
        while story is not None:
            for i in range(story.Fields.Count, 0, -1):   # reverse: Unlink removes fields
                field = story.Fields(i)
                parts = field.Code.Text.split()
                if len(parts) >= 2 and parts[0] == "MERGEFIELD" and parts[1] in values:
                    field.Result.Text = values[parts[1]]
                    field.Unlink()             # turn it into plain text
            story = story.NextStoryRange

    # 17 = PDF, check this before printing
    doc.SaveAs2(pdf_preview, FileFormat=17)
    print("Preview saved:", pdf_preview)

    if input("Print the form? (y/n): ").strip().lower() == "y":
        doc.PrintOut()
finally:
    doc.Close(SaveChanges=0)                   # never saves the template
    word.Quit()
