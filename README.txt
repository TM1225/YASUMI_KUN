Outlook CSV Scheduler
=====================

Overview
--------
This tool loads a UTF-8 CSV file that describes Outlook appointments,
displays the entries for review, and writes them to the default Outlook
desktop calendar through the COM API. The UI is implemented with Tkinter
and is intended to be distributed as a single EXE built with PyInstaller.

CSV Format
----------
Columns are fixed and the header row is mandatory.

    Date,Start,End,Subject,Status,Location,Body

Example values:

    2025-10-28,09:00,18:00,私用休暇,休み,,終日不在
    2025-10-29,13:00,14:00,通院,外出,クリニック,定期検査

Status Strings
--------------
Status values are mapped to Outlook's BusyStatus numeric values.

    休み / OOO / 不在 -> 3 (OutOfOffice)
    外出 / 忙しい     -> 2 (Busy)
    仮               -> 1 (Tentative)
    在席 / 空き       -> 0 (Free)
    他所勤務         -> 4 (WorkingElsewhere)
    Other strings    -> 2 (Busy, default)

Usage (Developer Runtime)
-------------------------
1. Ensure the host has:
   - Windows with Outlook desktop installed and configured.
   - Python 3.10+ and pywin32 (win32com).
2. Install dependencies if needed:

       pip install pywin32

3. Launch the app:

       python main.py

4. In the UI:
   - Click "CSVを選択" and pick your schedule CSV.
   - Review the preview grid.
   - Click "予定を登録" to push the entries to Outlook.
   - Monitor the log area for [OK] / [NG] per row.

Building a Standalone EXE
-------------------------
1. Install PyInstaller:

       pip install pyinstaller

2. Build the executable:

       pyinstaller --onefile --windowed main.py --name outlook-scheduler

3. The generated EXE is located in the `dist` directory as
   `outlook-scheduler.exe`. Distribute it together with
   `schedule_template.csv` and this `README.txt`.

Notes
-----
- The app uses the system timezone. Provide local times in the CSV.
- Outlook writes to the default calendar of the current profile.
- Rows with invalid dates/times are skipped; processing continues for the
  remaining rows.
- The first launch of the PyInstaller onefile EXE may take longer because
  the payload is extracted to a temporary directory.
