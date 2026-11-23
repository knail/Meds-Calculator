# Medication Supply Calculator

**Author:** KNail  
**Website:** [my02dcents.com](https://www.my02dcents.com/)  
**Version:** 1.0  
**Creation date:** 2026-11-22  
**Last updated:** 2026-11-22

A Windows desktop tool for estimating medication supply, checking remaining quantities on a shared date, viewing a monthly availability calendar, and preparing refill reminders for Outlook or another iCalendar-compatible calendar.

> **Important:** This is an estimate based on the information entered. It is not medical advice and does not replace prescription instructions. Confirm refill timing and dosing with your pharmacist or healthcare professional.

## Free use and disclaimer

This software is free to use and is provided **"AS IS"**, without warranties of any kind, express or implied, including merchantability, fitness for a particular purpose, and noninfringement. Use it at your own risk. To the extent permitted by law, the author is not liable for any claims, damages, or losses arising from its use, including incorrect estimates, data loss, or injury.

## Features

- Track **1–10 medications** in one session.
- Record each medication's optional name, fill date, quantity, and schedule.
- Estimate remaining quantity and the last day covered by each fill.
- Choose a shared spot-check date and see every medication's remaining supply on that date.
- View a monthly calendar with each medication's remaining quantity shown in each day's cell.
- Create one `.ics` file per refill date, combining medications that share the date in one appointment.
- Use a light gray interface with amber banners and a highlighted calendar button. Calendar entries use green for available supply, red for depleted supply, and gray before the first fill date.
- Use the Windows Segoe UI font and subtly tinted input fields for readability.

## Requirements

- Windows with an interactive desktop.
- Windows PowerShell 5.1.
- .NET Windows Forms and Drawing assemblies, which are loaded by the script.
- No additional PowerShell modules or package installation.

## Run the application

Open PowerShell in the folder containing the script and run:

```powershell
.\Meds-Calculator.ps1
```

The script opens the **Medication Supply Calculator** window.

## Portable Windows EXE

The calculator can also be packaged as a portable EXE using
[PS2EXE](https://github.com/MScholtes/PS2EXE). The EXE contains the PowerShell
script and runs it using Windows PowerShell and .NET Framework; it is not a
native Rust application. The target workstation needs Windows PowerShell 5.1,
.NET Framework 4.x, and an interactive desktop, but does not need PS2EXE.
Copy `Meds-Calculator.exe` to the workstation and double-click it to launch.
No installer or administrator privileges are required by the application.

To build the EXE, open Windows PowerShell 5.1 in the project folder and run:

```powershell
Install-Module -Name ps2exe -RequiredVersion 1.0.18 -Scope CurrentUser
Import-Module ps2exe -RequiredVersion 1.0.18
Invoke-PS2EXE -InputFile .\Meds-Calculator.ps1 -OutputFile .\Meds-Calculator.exe `
    -IconFile .\Meds-Calculator.ico -NoConsole -STA -Title 'Medication Supply Calculator' `
    -Product 'Medication Supply Calculator' -Version '1.0.0.0'
```

Installing PS2EXE requires internet access and may prompt for repository trust.
Rebuild the EXE after changing the script. An SSL certificate is not needed to
build or run it. The resulting EXE is unsigned, so Windows SmartScreen,
antivirus, or organizational application-control policies may warn about or
block it. A code-signing certificate is separate from an SSL certificate; check
with your administrator if workplace policy requires signed software.
If an older icon appears, check that the desktop copy or shortcut points to
the updated EXE. Windows may cache icons; refreshing Explorer or copying the
EXE under a new filename can help.

## Enter medication details

1. Select the number of medications to track, from 1 to 10.
2. For each medication, enter a name if desired. Blank names are labeled `Medication 1`, `Medication 2`, and so on in results and calendar events.
3. Enter the date the medication was filled and the quantity supplied.
4. Choose the schedule:
   - **Daily:** enter 1–4 doses per day.
   - **Weekly:** the calculator uses one dose every seven days, starting on the fill-date weekday. The daily dose control is set to one and disabled.
5. Set the shared **Spot check date**.
6. Select **Calculate All Medications** to display the remaining quantity and refill date for each entry.

The results table shows each medication name, its quantity on the selected
spot-check date, and its refill date. Scroll for additional medications or
hover over a name to see it in full.

Reducing the medication count hides later entries but keeps their values in the session. Increasing the count again restores them.

To copy Medication 1's fill date to the other currently displayed medications,
select **Copy Med 1 Fill Date to All**. This is a one-time copy: later changes
or newly added medications are not automatically synchronized, and hidden
entries are unchanged. Existing results are recalculated after the copy.

## How supply estimates work

Quantities are measured **at the start of the selected date**. Doses scheduled on earlier dates are deducted; a dose scheduled for the selected date is not deducted yet.

### Daily schedule

The refill date is the last day covered by the supply:

```text
refill date = fill date + ceil(quantity / doses per day) - 1 day
```

For example, 30 units taken twice daily cover 15 days beginning on the fill date. The refill date is the fifteenth covered day.

### Weekly schedule

The calculator assumes one dose on the fill-date weekday and one dose every seven days afterward:

```text
refill date = fill date + (quantity - 1) * 7 days
```

For both schedules, if the spot-check date is before the fill date, the result is shown as **not filled yet**.

## Availability calendar

Select **View Availability Calendar** to open the month containing the shared spot-check date. Use **Previous** and **Next** to navigate.

- **Green:** this medication has supply.
- **Red:** this medication's supply is depleted.
- **Gray:** this medication has not reached its fill date.

Each day cell lists every medication as `Name | remaining quantity`. A medication that has not reached its fill date is labeled **Not filled**; a depleted supply shows `0`. The shared spot-check date is outlined, and its values are measured at the start of that date. Larger calendar text may leave less room for entries; hover over a day to see its complete list if they do not all fit in the cell.

## Create refill appointments

In **Create Refill Appointments**, choose the appointment time and enter how
many days before each refill date to schedule it. Use `0` for the refill date
itself. The default time is 9:00 AM. Select **Create Refill Appointments (.ics)**
and choose a folder:

- One `.ics` file is created for each distinct refill date. Medications with the same refill date are combined in one appointment, with all their names in its title.
- Each appointment occurs on its calculated refill date or the selected number of days before it.
- Each appointment lasts one hour and uses the shared appointment time.
- Each appointment includes an alert one day before the appointment, even when the appointment is scheduled before the refill date.
- Each event's description includes the actual refill date.
- Filenames include the refill date. Existing files are not overwritten; a numeric suffix is added when needed.
- The folder picker starts in Documents when that folder is available.
- After saving, the script attempts to open each file with the system's default `.ics` application. Outlook can import or open the files when it is configured as the calendar handler.

The appointment time is interpreted in the computer's local time zone and converted for the calendar file. The script reports an error if the selected time does not exist on an appointment date because of a daylight-saving transition, or if the offset would produce an invalid date.

## Tests

Run the focused calendar-status and appointment-generation tests in Windows PowerShell:

```powershell
.\tests\Meds-Calculator.Tests.ps1
```

## Calculation limits

The estimates assume a regular schedule and the entered quantity. The calculator does not model:

- As-needed or irregular use.
- Missed or extra doses.
- Dose changes, tapering, or changing schedules.
- Partial fills, early refills, or quantities carried over from a previous fill.

Review the entered details and use the prescription label and professional guidance when deciding when to request or take medication.

## PowerShell file organization

The script begins with PowerShell comment-based help, followed by foldable sections for:

1. Application theme and form setup.
2. Main window and medication input controls.
3. Shared date, result, and action controls.
4. Supply calculations and iCalendar text helpers.
5. Medication result calculations.
6. Monthly medication availability calendar.
7. iCalendar generation and application event handlers.

In PowerShell, the script help can be viewed with:

```powershell
Get-Help .\Meds-Calculator.ps1 -Full
```
