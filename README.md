# label-mailmerge-automation
"Desktop tool that automates label and mail-merge form printing for a school district's damaged-device replacement workflow (tkinter + Brother b-PAC SDK + Word mail merge via pywin32)."

READ ME
A Python/tkinter desktop app that automates the repetitive parts of processing a student damaged-device replacement: entering device info once, printing a Brother label via the b-PAC SDK, and printing a single mail-merge form via Word/COM — pulling from a shared Excel database instead of manually loading the same spreadsheet twice in two different programs. Config-driven so it isn't tied to any one organization's fields or file paths.
