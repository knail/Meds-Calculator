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
- Create one `.ics` file containing a separate refill appointment for each medication.
- Use a light gray interface with deep teal accents. Calendar statuses use green for available supply, red for depleted supply, and gray before the first fill date.
- Use a modern system font when available, with a fallback for other Windows installations.

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

## Enter medication details

1. Select the number of medications to track, from 1 to 10.
2. For each medication, enter a name if desired. Blank names are labeled `Medication 1`, `Medication 2`, and so on in results and calendar events.
3. Enter the date the medication was filled and the quantity supplied.
4. Choose the schedule:
   - **Daily:** enter 1–4 doses per day.
   - **Weekly:** the calculator uses one dose every seven days, starting on the fill-date weekday. The daily dose control is set to one and disabled.
5. Set the shared **Spot check date** and, if you plan to export appointments, the shared **Appointment time**. The appointment time defaults to 9:00 AM.
6. Select **Calculate all medications** to display the remaining quantity and refill date for each entry.

Reducing the medication count hides later entries but keeps their values in the session. Increasing the count again restores them.

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

Select **View availability calendar** to open the month containing the shared spot-check date. Use **Previous** and **Next** to navigate.

- **Green:** at least one medication that has been filled still has supply.
- **Red:** at least one medication has been filled, and all filled supplies are depleted.
- **Gray:** no medication has reached its fill date yet.

Each day cell lists every medication as `Name | remaining quantity`. A medication that has not reached its fill date is labeled **Not filled**; a depleted supply shows `0`. The shared spot-check date is outlined, and its values are measured at the start of that date. Larger calendar text may leave less room for entries; hover over a day to see its complete list if they do not all fit in the cell.

## Create refill appointments

Select **Create refill appointments (.ics)** to create one combined calendar file:

- Each configured medication receives its own appointment on its calculated refill date.
- Each appointment lasts one hour and uses the shared appointment time.
- Each appointment includes a reminder one day before it starts.
- The suggested filename includes the medication names; unsafe filename characters are replaced, and long names are shortened.
- The save dialog starts in Documents when that folder is available.
- After saving, the script attempts to open the file with the system's default `.ics` application. Outlook can import or open the file when it is configured as the calendar handler.

The appointment time is interpreted in the computer's local time zone and converted for the calendar file. The script reports an error if the selected time does not exist on a refill date because of a daylight-saving transition.

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
