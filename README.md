# Work Achievements Logger

A lightweight Windows desktop application for recording completed work throughout the day.

## Features

- Required title and description fields.
- Automatic local completion date/time.
- Daily, weekly, and monthly views.
- Edit and delete existing entries.
- Automatic Excel-compatible `.xlsx` workbook updates.
- Export the selected view to a readable Word `.docx` report.
- No internet connection or account is required.
- Uses SQLite as the reliable local source of truth and mirrors entries to Excel.

## Requirements

- Windows 11 (tested design target: Enterprise 25H2).
- Python 3.11 or newer for running from source.
- Microsoft Excel is not required; the workbook can be opened with Excel or another compatible spreadsheet application.

Install dependencies:

```powershell
py -m pip install -r requirements.txt
```

## Run

```powershell
py app.py
```

The application stores data in:

```text
%USERPROFILE%\Documents\Work Achievements\work_achievements.db
%USERPROFILE%\Documents\Work Achievements\work_achievements.xlsx
```

The Excel workbook is regenerated after additions, edits, and deletions. The workbook contains one `Entries` worksheet with columns for completion date/time, title, description, and the calculated week/month labels.

## Daily, weekly, and monthly organization

Use the view selector to choose `Daily`, `Weekly`, or `Monthly`:

- **Daily**: select a date and see entries completed that day.
- **Weekly**: select any date and see the Monday-through-Sunday week containing it.
- **Monthly**: select any date and see all entries in that calendar month.

The selected view can be exported to Word with the **Export current view to Word** button.

## Build a standalone executable

Install the packaging dependency and run the build script:

```powershell
py -m pip install -r requirements-build.txt
.\build.ps1
```

The executable will be placed in `dist\WorkAchievementsLogger.exe`. PyInstaller bundles Python and the dependencies, so Python is not required on the target computer.

## Tests

```powershell
py -m pytest
```

## Notes

- SQLite is used internally to avoid losing data if Excel is closed or unavailable.
- Excel and Word exports are written to the same Work Achievements folder.
- If an export fails, the existing database remains intact and an error is shown in the application.
