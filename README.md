# Device Replacement Tool

A small Windows desktop tool that automates the repetitive end of a school district's damaged-device replacement workflow. After a replacement device is assigned, one form captures the new device's details and then:

1. **Prints an asset label** on a Brother label printer (via the Brother b-PAC SDK)
2. **Prints the student device agreement form** from a Word template (via Word automation)
3. **Logs the assignment** as a new row in an Excel spreadsheet

Previously these were three manual steps across three different applications (P-touch Editor, Word mail merge, and Excel).

---

## Requirements

This tool runs on **Windows only**. The exe bundles the Python code, but **not** the following. Each must be installed on the machine separately:

| Requirement                    | Notes                                                                                                                                                                                                                                   |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Brother b-PAC SDK (64-bit)** | Required for label printing. This is a **separate installer from P-touch Editor**: installing Editor alone does not provide it. It **must be the 64-bit version** to match the exe; the 32-bit SDK fails the same way as a missing one. |
| **Microsoft Word**             | Required for printing the form.                                                                                                                                                                                                         |
| **Printer drivers**            | For the Brother label printer and the regular (form) printer.                                                                                                                                                                           |

If b-PAC is missing, the tool still opens with a warning; label printing is unavailable, while form printing and Excel logging still work.

---

## Installation

1. Create a permanent folder for the tool, e.g. `C:\Tools\DeviceReplacementTool\`.
2. Copy `DeviceReplacementTool.exe` into it.
3. Copy `config.example.toml` into the same folder, rename it to `config.toml`, and fill in your paths (see below).
4. To launch from the desktop or taskbar, right-click the exe → **Create shortcut**, or pin it to the taskbar.

> **Keep `config.toml` in the same folder as the exe.** The tool looks for it next to itself. Move the exe on its own and it will report that the config is missing — use a shortcut instead.

---

## Configuration (`config.toml`)

```toml
[paths]
excel_database = 'C:\path\to\Database.xlsx'
label_template = 'C:\path\to\Device Label.lbx'
form_template  = 'C:\path\to\Device Agreement.docx'

[schools]   # display name = school code; order here = dropdown order
"Example High School" = "EHS"
"Example Middle School" = "EMS"

[devices]
models = ["Example Chromebook Model A", "Example Chromebook Model B"]
```

**Editing:** the file is plain text, so Notepad works fine.

- **Windows paths must use single quotes** (`'C:\...'`). Double quotes treat backslashes as escape characters.
- Save as **UTF-8** (Notepad's default), not "UTF-8 with BOM".
- When using _Save As_ in Notepad, set _Save as type_ to **All files**, or it may save as `config.toml.txt`. Turn on _View → Show → File name extensions_ in Explorer to spot this.
- If the config is missing or has an error, the tool shows a popup explaining what's wrong and closes.

Paths are machine-specific: when setting up another computer, point them at wherever the templates and database live on that machine.

---

## Template setup

The tool fills in fields directly in each template. It does not use a real mail merge.

### Label template (`.lbx`, created in P-touch Editor)

The template must contain objects with these exact names:

| Object name | Filled with          |
| ----------- | -------------------- |
| `Name`      | Student name         |
| `School`    | School               |
| `YOG`       | Year of graduation   |
| `Bar Code`  | Device serial number |
| `QR Code`   | Device serial number |

Other objects (static captions, logos) are left untouched. The printer is chosen in the tool at print time, so the printer saved in the template doesn't matter and the template is never modified.

### Form template (`.docx`)

- Must be a **Normal Word Document**, not a mail-merge document with an attached data source.
- Uses MERGEFIELDs named `User`, `Grade`, `Device`, `Serial`, and `SchoolCode` (anywhere in the document, including headers and footers).
- `SchoolCode` is filled with the school's code from the `[schools]` section of the config.
- The template is opened **read-only** and is never saved or changed.

### Excel database (`.xlsx`)

- Must be `.xlsx` (older `.xls` files are not supported).
- Each submission appends a row with columns in this order: **Serial, User, YOG, Grade, School, Device**.
- If new rows appear with a gap above them, there are formatted-but-empty rows at the bottom of the sheet: delete those rows in Excel.

---

## Usage

1. Enter the device serial number, student name, year of graduation, and grade.
2. Choose the school and device model.
3. Choose the label printer and form printer.
4. Tick the actions to perform (all on by default):
   - **Print Labels**
   - **Print Forms**
   - **Print Database** (saves the row to the Excel spreadsheet)
5. Click **Submit**.

Actions run in the order **label → form → Excel**. If any step fails, the tool stops there and shows an error, so nothing after the failed step runs. After a fully successful submit, all three checkboxes reset to on.

### Recovering from a partial run

Use the checkboxes to redo only what's needed. For example, if the label and form printed but the spreadsheet was locked:

1. Close the spreadsheet (or wait for OneDrive to finish syncing).
2. Uncheck **Print Labels** and **Print Forms**.
3. Submit again to save just the row.

The same approach works for reprinting just a label or just a form.

---

## Troubleshooting

| Problem                                            | Fix                                                                                                                                                   |
| -------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| "b-PAC SDK could not be loaded" warning at startup | Install the **64-bit** Brother b-PAC SDK, then restart the tool.                                                                                      |
| Label printer list is empty                        | Check that the Brother printer is installed and b-PAC is set up, then restart the tool.                                                               |
| "Config Missing" / "Config Error" popup            | Make sure `config.toml` is next to the exe, paths are in single quotes, and all settings from `config.example.toml` are present.                      |
| "Excel Error" — spreadsheet open or locked         | Close the file in Excel (or let OneDrive finish syncing), then resubmit with only **Print Database** checked.                                         |
| Nothing prints, but no error appears               | The printer is probably offline. Jobs usually just sit in the Windows print queue without an error, so check the printer and its queue.               |
| Windows default printer changed                    | It shouldn't: the tool restores the default printer after printing a form. If Word was force-closed mid-print, reset the default in Windows settings. |
| Antivirus quarantines the exe                      | Single-file PyInstaller exes are sometimes flagged. Build with `--onedir` instead (see Building).                                                     |

---

## Development

### Setup

Requires Python 3.14 (64-bit).

```
py -3.14 -m pip install -r requirements-dev.txt
```

- `requirements.txt`: runtime dependencies (`openpyxl`, `pywin32`)
- `requirements-dev.txt`: runtime dependencies plus `pyinstaller`

Run the script directly while developing:

```
py -3.14 label_form_automation.py
```

Copy `config.example.toml` to `config.toml` next to the script first. The real `config.toml` is git-ignored and should never be committed.

### Building the exe

```
py -3.14 -m PyInstaller DeviceReplacementTool.spec
```

The exe is written to `dist\`. Copy a `config.toml` next to it to test. The config is intentionally **not** bundled into the exe, so keep `datas=[]` empty in the spec.

**If the exe closes silently:** it's built without a console, so errors are invisible. Rebuild without `--windowed` (set `console=True` in the spec) and run the exe from a terminal to see the traceback. Missing modules can be added as hidden imports (e.g. `pywintypes`, `win32timezone`).

**Single-file vs. folder:** the default single-file build starts a little slowly because it unpacks itself on each launch. That's normal. If antivirus flags it, a `--onedir` build (a folder with the exe and its DLLs) is the fallback and needs no code changes.

### Release routine

1. Finish and test changes by running the script.
2. Commit and merge to `main`.
3. Rebuild from the spec and quickly test the new exe (launch + one submit).
4. Replace the deployed exe. `config.toml` stays where it is.
5. Tag the release (`git tag v1.1`).

### Technical notes

**b-PAC uses late binding.** The tool uses `win32com.client.Dispatch("bpac.Document")`. `gencache.EnsureDispatch` does not work with b-PAC. A side effect of late binding:

- Methods that take **no arguments** are called when accessed, so they are written **without parentheses** (`doc.Close`, `doc.EndPrint`, `doc.Printer.GetInstalledPrinters`). Adding `()` causes `TypeError: 'bool'/'tuple' object is not callable`.
- Methods with arguments keep their parentheses (`doc.Open(path)`, `doc.PrintOut(1, 0)`).
- Properties never use parentheses.

**b-PAC `Export` file types.** Older Brother documentation lists these incorrectly. The correct values are: 0 = same as opened, 1 = LBX, 2 = LBL, 3 = LBI, **4 = BMP**, 5 = PAF. For example, `doc.Export(4, path, 180)` saves a BMP preview.

**Word and the default printer.** Setting `word.ActivePrinter` changes the Windows default printer, and `PrintOut` has no printer argument. The tool saves the original default, prints, and restores it afterwards.

**Stale pywin32 cache.** If you get odd `AttributeError`s after updating pywin32 or reinstalling Word/b-PAC, delete the `gen_py` folder (its location is `win32com.__gen_path__`) and run again.

---

## Roadmap

- Pick an existing database or create a new one from within the tool
- Convert older `.xls` databases to `.xlsx`
- Remember the last-used printers
- Bulk mode: print labels and forms for every row in a spreadsheet
