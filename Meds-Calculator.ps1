<#
.SYNOPSIS
Estimates medication supply and creates refill calendar appointments.

.DESCRIPTION
Medication Supply Calculator is a Windows Forms application for tracking up to
10 medications. It estimates remaining quantities and refill dates from each
medication's fill date, quantity, and daily or weekly schedule. It can also
show a monthly availability overview and create Outlook-compatible .ics files
with one one-hour, one-day-reminder event per refill date.

.NOTES
Script Name:   Meds-Calculator.ps1
Author:        KNail
Website:       https://www.my02dcents.com/
Version:       1.0
Creation Date: 2026-11-22
Last Updated:  2026-11-22
Purpose:       Estimate medication supply and prepare refill reminders.

This software is free to use and is provided "AS IS", without warranties of
any kind, express or implied, including merchantability, fitness for a
particular purpose, and noninfringement. Use it at your own risk. To the
extent permitted by law, the author is not liable for any claims, damages,
or losses arising from its use, including incorrect estimates, data loss,
or injury.

Supply estimates are measured at the start of the selected date. The daily
schedule supports 1-4 doses per day. Weekly means one dose every seven days,
starting on the fill-date weekday. This tool does not account for missed doses,
as-needed use, dose changes, partial fills, or other irregular schedules.
It is not medical advice and does not replace prescription instructions.
Verify estimates and refill instructions with a qualified healthcare professional.

.EXAMPLE
.\Meds-Calculator.ps1
Launches the medication supply calculator.

.LINK
https://www.my02dcents.com/
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Application theme and form setup
$Script:MaxMedicationCount = 10
$Script:MedicationEntries = [System.Collections.Generic.List[System.Object]]::new()
$Script:CurrentMedicationResults = @()
$Script:ThemePrimary = [System.Drawing.Color]::FromArgb(0, 100, 75)
$Script:ThemeBanner = [System.Drawing.Color]::FromArgb(170, 90, 0)
$Script:ThemeBannerHover = [System.Drawing.Color]::FromArgb(136, 72, 0)
$Script:ThemeAccent = [System.Drawing.Color]::FromArgb(0, 108, 115)
$Script:ThemeText = [System.Drawing.Color]::FromArgb(43, 52, 59)
$Script:ThemeMutedText = [System.Drawing.Color]::FromArgb(91, 105, 112)
$Script:ThemeBackground = [System.Drawing.Color]::FromArgb(238, 242, 244)
$Script:ThemeSurface = [System.Drawing.Color]::FromArgb(255, 255, 255)
$Script:ThemeSurfaceAlt = [System.Drawing.Color]::FromArgb(247, 249, 250)
$Script:ThemeInput = [System.Drawing.Color]::FromArgb(242, 248, 244)
$Script:ThemeBorder = [System.Drawing.Color]::FromArgb(192, 203, 208)
$Script:ThemeHoverSurface = [System.Drawing.Color]::FromArgb(231, 241, 242)
$Script:ThemeGreen = [System.Drawing.Color]::FromArgb(226, 243, 232)
$Script:ThemeGreenText = [System.Drawing.Color]::FromArgb(37, 105, 66)
$Script:ThemeRedText = [System.Drawing.Color]::FromArgb(151, 59, 54)
$Script:UiFontFamily = 'Segoe UI'
$Script:DatePickerFont = New-Object System.Drawing.Font($Script:UiFontFamily, 10)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Medication Calculator'
$form.Size = New-Object System.Drawing.Size(780, [Math]::Min(835, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.AutoScroll = $true
$form.AutoScrollMargin = New-Object System.Drawing.Size(0, 10)
$form.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10)
$form.BackColor = $Script:ThemeBackground
$form.ForeColor = $Script:ThemeText

$applicationIcon = $null
# PS2EXE runs the embedded script without a script root.
if ([string]::IsNullOrEmpty($PSScriptRoot))
{
    $applicationIcon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    if ($null -eq $applicationIcon)
    {
        throw 'Could not load the application icon from the executable.'
    }
}
else
{
    $iconPath = Join-Path $PSScriptRoot 'Meds-Calculator.ico'
    if (Test-Path -LiteralPath $iconPath -PathType Leaf)
    {
        $applicationIcon = New-Object System.Drawing.Icon($iconPath)
    }
}
if ($null -ne $applicationIcon)
{
    $form.Icon = $applicationIcon
}

function New-SectionCard([string]$Title, [int]$X, [int]$Y, [int]$Width, [int]$Height) {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point($X, $Y)
    $panel.Size = New-Object System.Drawing.Size($Width, $Height)
    $panel.BackColor = $Script:ThemeSurface
    $panel.ForeColor = $Script:ThemeText
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = $Title
    $heading.Location = New-Object System.Drawing.Point(14, 7)
    $heading.Size = New-Object System.Drawing.Size(($Width - 28), 22)
    $heading.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10, [System.Drawing.FontStyle]::Bold)
    $heading.ForeColor = [System.Drawing.Color]::Black
    $panel.Controls.Add($heading)
    $form.Controls.Add($panel)
    return $panel
}

function Add-FieldLabel([System.Windows.Forms.Control]$Parent, [string]$Text, [int]$X, [int]$Y, [int]$Width) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, ($Y + 3))
    $label.Size = New-Object System.Drawing.Size($Width, 22)
    $label.ForeColor = $Script:ThemeText
    $Parent.Controls.Add($label)
}

function Set-ButtonStyle([System.Windows.Forms.Button]$Button, [bool]$Primary = $false, [float]$FontSize = 10) {
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatAppearance.BorderSize = if ($Primary) { 0 } else { 1 }
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-Object System.Drawing.Font($Script:UiFontFamily, $FontSize, [System.Drawing.FontStyle]::Bold)
    if ($Primary)
    {
        $Button.BackColor = $Script:ThemeBanner
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatAppearance.MouseOverBackColor = $Script:ThemeBannerHover
    }
    else
    {
        $Button.BackColor = $Script:ThemeSurface
        $Button.ForeColor = $Script:ThemeText
        $Button.FlatAppearance.BorderColor = $Script:ThemeBorder
        $Button.FlatAppearance.MouseOverBackColor = $Script:ThemeHoverSurface
    }
}

#endregion

#region Main window and medication input controls
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(764, 70)
$headerPanel.BackColor = $Script:ThemeBanner
$headerPanel.ForeColor = $Script:ThemeText
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = 'Medication Supply Calculator'
$titleLabel.Location = New-Object System.Drawing.Point(18, 7)
$titleLabel.Size = New-Object System.Drawing.Size(700, 29)
$titleLabel.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$headerPanel.Controls.Add($titleLabel)

$instructionLabel = New-Object System.Windows.Forms.Label
$instructionLabel.Text = 'Add each medication, then check all supplies for one shared date.'
$instructionLabel.Location = New-Object System.Drawing.Point(20, 40)
$instructionLabel.Size = New-Object System.Drawing.Size(700, 21)
$instructionLabel.ForeColor = [System.Drawing.Color]::White
$headerPanel.Controls.Add($instructionLabel)

$countPanel = New-Object System.Windows.Forms.Panel
$countPanel.Location = New-Object System.Drawing.Point(16, 78)
$countPanel.Size = New-Object System.Drawing.Size(712, 44)
$countPanel.BackColor = $Script:ThemeSurface
$countPanel.ForeColor = $Script:ThemeText
$countPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($countPanel)

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Text = 'Number of medications:'
$countLabel.Location = New-Object System.Drawing.Point(12, 10)
$countLabel.Size = New-Object System.Drawing.Size(165, 22)
$countLabel.ForeColor = $Script:ThemeText
$countPanel.Controls.Add($countLabel)

$medicationCountInput = New-Object System.Windows.Forms.NumericUpDown
$medicationCountInput.Minimum = 1
$medicationCountInput.Maximum = $Script:MaxMedicationCount
$medicationCountInput.Value = 1
$medicationCountInput.Location = New-Object System.Drawing.Point(180, 8)
$medicationCountInput.Size = New-Object System.Drawing.Size(70, 26)
$medicationCountInput.BackColor = $Script:ThemeInput
$medicationCountInput.ForeColor = $Script:ThemeText
$countPanel.Controls.Add($medicationCountInput)

$countHint = New-Object System.Windows.Forms.Label
$countHint.Text = "Choose 1-$($Script:MaxMedicationCount)"
$countHint.Location = New-Object System.Drawing.Point(265, 11)
$countHint.Size = New-Object System.Drawing.Size(90, 22)
$countHint.ForeColor = $Script:ThemeMutedText
$countPanel.Controls.Add($countHint)

$copyFillDateButton = New-Object System.Windows.Forms.Button
$copyFillDateButton.Text = 'Copy Med 1 Fill Date to All'
$copyFillDateButton.Location = New-Object System.Drawing.Point(365, 6)
$copyFillDateButton.Size = New-Object System.Drawing.Size(330, 30)
Set-ButtonStyle -Button $copyFillDateButton -FontSize 9
$countPanel.Controls.Add($copyFillDateButton)

$medicationListPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$medicationListPanel.Location = New-Object System.Drawing.Point(16, 132)
$medicationListPanel.Size = New-Object System.Drawing.Size(712, 228)
$medicationListPanel.AutoScroll = $true
$medicationListPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$medicationListPanel.WrapContents = $false
$medicationListPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$medicationListPanel.BackColor = $Script:ThemeSurfaceAlt
$medicationListPanel.ForeColor = $Script:ThemeText
$medicationListPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($medicationListPanel)

$medicationToolTip = New-Object System.Windows.Forms.ToolTip
$medicationToolTip.SetToolTip($copyFillDateButton, 'Copies Medication 1''s fill date to the currently displayed medications. Hidden entries are unchanged.')

function New-MedicationEntry([int]$Number) {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(680, 100)
    $panel.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 5)
    $panel.BackColor = $Script:ThemeSurface
    $panel.ForeColor = $Script:ThemeText
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = "MEDICATION $Number"
    $heading.Location = New-Object System.Drawing.Point(10, 5)
    $heading.Size = New-Object System.Drawing.Size(650, 20)
    $heading.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
    $heading.ForeColor = [System.Drawing.Color]::Black
    $panel.Controls.Add($heading)

    Add-FieldLabel -Parent $panel -Text 'Name:' -X 12 -Y 30 -Width 100
    $nameInput = New-Object System.Windows.Forms.TextBox
    $nameInput.Location = New-Object System.Drawing.Point(120, 28)
    $nameInput.Size = New-Object System.Drawing.Size(235, 25)
    $nameInput.MaxLength = 100
    $nameInput.BackColor = $Script:ThemeInput
    $nameInput.ForeColor = $Script:ThemeText
    $panel.Controls.Add($nameInput)
    $medicationToolTip.SetToolTip($nameInput, 'Optional. A generic medication label is used when blank.')

    Add-FieldLabel -Parent $panel -Text 'Fill Date:' -X 360 -Y 30 -Width 80
    $fillDatePicker = New-Object System.Windows.Forms.DateTimePicker
    $fillDatePicker.Format = 'Short'
    $fillDatePicker.Location = New-Object System.Drawing.Point(445, 28)
    $fillDatePicker.AutoSize = $false
    $fillDatePicker.Size = New-Object System.Drawing.Size(227, 28)
    $fillDatePicker.Font = $Script:DatePickerFont
    $fillDatePicker.CalendarMonthBackground = $Script:ThemeInput
    $fillDatePicker.CalendarForeColor = $Script:ThemeText
    $panel.Controls.Add($fillDatePicker)

    Add-FieldLabel -Parent $panel -Text 'Quantity:' -X 12 -Y 69 -Width 100
    $quantityInput = New-Object System.Windows.Forms.NumericUpDown
    $quantityInput.Minimum = 1
    $quantityInput.Maximum = 10000
    $quantityInput.Value = 30
    $quantityInput.Location = New-Object System.Drawing.Point(120, 67)
    $quantityInput.Size = New-Object System.Drawing.Size(100, 25)
    $quantityInput.BackColor = $Script:ThemeInput
    $quantityInput.ForeColor = $Script:ThemeText
    $panel.Controls.Add($quantityInput)

    Add-FieldLabel -Parent $panel -Text 'Doses taken:' -X 230 -Y 69 -Width 90
    $dosesInput = New-Object System.Windows.Forms.NumericUpDown
    $dosesInput.Minimum = 1
    $dosesInput.Maximum = 4
    $dosesInput.Value = 1
    $dosesInput.Location = New-Object System.Drawing.Point(325, 67)
    $dosesInput.Size = New-Object System.Drawing.Size(65, 25)
    $dosesInput.BackColor = $Script:ThemeInput
    $dosesInput.ForeColor = $Script:ThemeText
    $panel.Controls.Add($dosesInput)

    $frequencyBox = New-Object System.Windows.Forms.ComboBox
    $frequencyBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$frequencyBox.Items.AddRange([object[]]@('Daily', 'Weekly'))
    $frequencyBox.SelectedIndex = 0
    $frequencyBox.Location = New-Object System.Drawing.Point(400, 67)
    $frequencyBox.Size = New-Object System.Drawing.Size(130, 25)
    $frequencyBox.BackColor = $Script:ThemeInput
    $frequencyBox.ForeColor = $Script:ThemeText
    $panel.Controls.Add($frequencyBox)
    $medicationToolTip.SetToolTip($frequencyBox, 'Weekly means one dose on the fill-date weekday, repeated every 7 days.')
    $frequencyBox.Add_SelectedIndexChanged({
        if ($frequencyBox.SelectedItem -eq 'Weekly')
        {
            $dosesInput.Value = 1
            $dosesInput.Enabled = $false
        }
        else
        {
            $dosesInput.Enabled = $true
        }
    }.GetNewClosure())

    return [PSCustomObject]@{
        Panel = $panel
        Heading = $heading
        NameInput = $nameInput
        FillDatePicker = $fillDatePicker
        QuantityInput = $quantityInput
        DosesInput = $dosesInput
        FrequencyBox = $frequencyBox
    }
}

function Sync-MedicationEntries {
    $targetCount = [int]$medicationCountInput.Value
    while ($Script:MedicationEntries.Count -lt $targetCount)
    {
        $entry = New-MedicationEntry -Number ($Script:MedicationEntries.Count + 1)
        [void]$Script:MedicationEntries.Add($entry)
        $medicationListPanel.Controls.Add($entry.Panel)
    }
    for ($index = 0; $index -lt $Script:MedicationEntries.Count; $index++)
    {
        $entry = $Script:MedicationEntries[$index]
        $entry.Panel.Visible = $index -lt $targetCount
    }
}
$medicationCountInput.Add_ValueChanged({ Sync-MedicationEntries })
Sync-MedicationEntries

#endregion

#region Shared date, result, and action controls
$sharedCard = New-SectionCard -Title 'Spot Check Date' -X 16 -Y 370 -Width 712 -Height 74
$spotHeading = $sharedCard.Controls[0]
$spotHeading.Location = New-Object System.Drawing.Point(170, 23)
$spotHeading.Size = New-Object System.Drawing.Size(180, 28)
$spotHeading.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$checkDatePicker = New-Object System.Windows.Forms.DateTimePicker
$checkDatePicker.Format = 'Short'
$checkDatePicker.Location = New-Object System.Drawing.Point(360, 23)
$checkDatePicker.AutoSize = $false
$checkDatePicker.Size = New-Object System.Drawing.Size(180, 28)
$checkDatePicker.Font = $Script:DatePickerFont
$checkDatePicker.CalendarMonthBackground = $Script:ThemeInput
$checkDatePicker.CalendarForeColor = $Script:ThemeText
$sharedCard.Controls.Add($checkDatePicker)

$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Text = 'Calculate All Medications'
$calculateButton.Location = New-Object System.Drawing.Point(16, 454)
$calculateButton.Size = New-Object System.Drawing.Size(712, 38)
$calculateButton.TabIndex = 1
Set-ButtonStyle -Button $calculateButton -Primary $true -FontSize 11
$form.Controls.Add($calculateButton)
$form.AcceptButton = $calculateButton

$resultPanel = New-SectionCard -Title 'Remaining Quantity and Refill Date' -X 16 -Y 502 -Width 712 -Height 110
$resultsList = New-Object System.Windows.Forms.ListView
$resultsList.View = [System.Windows.Forms.View]::Details
$resultsList.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::Nonclickable
$resultsList.FullRowSelect = $true
$resultsList.GridLines = $true
$resultsList.ShowItemToolTips = $true
$resultsList.BackColor = $Script:ThemeInput
$resultsList.ForeColor = $Script:ThemeText
$resultsList.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 9)
$resultsList.Location = New-Object System.Drawing.Point(14, 31)
$resultsList.Size = New-Object System.Drawing.Size(680, 68)
[void]$resultsList.Columns.Add('Med Name', 230)
[void]$resultsList.Columns.Add('Qty on Spot Check Date', 245)
[void]$resultsList.Columns.Add('Refill Date', 170)
$resultPanel.Controls.Add($resultsList)

$resultsHint = New-Object System.Windows.Forms.Label
$resultsHint.Text = 'Enter medication details, then calculate to see each remaining quantity.'
$resultsHint.Location = New-Object System.Drawing.Point(20, 54)
$resultsHint.Size = New-Object System.Drawing.Size(640, 20)
$resultsHint.BackColor = $Script:ThemeInput
$resultsHint.ForeColor = $Script:ThemeText
$resultsHint.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 9)
$resultPanel.Controls.Add($resultsHint)
$resultsHint.BringToFront()

$calendarButton = New-Object System.Windows.Forms.Button
$calendarButton.Text = 'View Availability Calendar'
$calendarButton.Location = New-Object System.Drawing.Point(16, 622)
$calendarButton.Size = New-Object System.Drawing.Size(712, 36)
$calendarButton.TabIndex = 2
Set-ButtonStyle -Button $calendarButton -FontSize 11
$calendarButton.BackColor = $Script:ThemeGreen
$calendarButton.ForeColor = $Script:ThemeGreenText
$calendarButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(207, 234, 215)
$form.Controls.Add($calendarButton)

$appointmentCard = New-SectionCard -Title 'Create Refill Appointments' -X 16 -Y 670 -Width 712 -Height 118
Add-FieldLabel -Parent $appointmentCard -Text 'Days before refill:' -X 12 -Y 36 -Width 130
$daysBeforeInput = New-Object System.Windows.Forms.NumericUpDown
$daysBeforeInput.Minimum = 0
$daysBeforeInput.Maximum = ([datetime]::MaxValue.Date - [datetime]::MinValue.Date).Days
$daysBeforeInput.Value = 0
$daysBeforeInput.Location = New-Object System.Drawing.Point(145, 33)
$daysBeforeInput.Size = New-Object System.Drawing.Size(90, 26)
$daysBeforeInput.BackColor = $Script:ThemeInput
$daysBeforeInput.ForeColor = $Script:ThemeText
$appointmentCard.Controls.Add($daysBeforeInput)

Add-FieldLabel -Parent $appointmentCard -Text 'Appointment time:' -X 255 -Y 36 -Width 125
$appointmentTimePicker = New-Object System.Windows.Forms.DateTimePicker
$appointmentTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$appointmentTimePicker.ShowUpDown = $true
$appointmentTimePicker.Value = (Get-Date).Date.AddHours(9)
$appointmentTimePicker.Location = New-Object System.Drawing.Point(383, 33)
$appointmentTimePicker.AutoSize = $false
$appointmentTimePicker.Size = New-Object System.Drawing.Size(120, 28)
$appointmentTimePicker.BackColor = $Script:ThemeInput
$appointmentTimePicker.ForeColor = $Script:ThemeText
$appointmentTimePicker.CalendarMonthBackground = $Script:ThemeInput
$appointmentTimePicker.CalendarForeColor = $Script:ThemeText
$appointmentCard.Controls.Add($appointmentTimePicker)

$appointmentDateHint = New-Object System.Windows.Forms.Label
$appointmentDateHint.Text = '0 = refill day'
$appointmentDateHint.Location = New-Object System.Drawing.Point(520, 36)
$appointmentDateHint.Size = New-Object System.Drawing.Size(175, 22)
$appointmentDateHint.ForeColor = $Script:ThemeMutedText
$appointmentCard.Controls.Add($appointmentDateHint)

$appointmentButton = New-Object System.Windows.Forms.Button
$appointmentButton.Text = 'Create Refill Appointments (.ics)'
$appointmentButton.Location = New-Object System.Drawing.Point(12, 72)
$appointmentButton.Size = New-Object System.Drawing.Size(430, 36)
$appointmentButton.TabIndex = 3
Set-ButtonStyle -Button $appointmentButton -FontSize 11
$appointmentCard.Controls.Add($appointmentButton)

$appointmentHint = New-Object System.Windows.Forms.Label
$appointmentHint.Text = '1 hr; alert 1 day before appointment'
$appointmentHint.Location = New-Object System.Drawing.Point(452, 79)
$appointmentHint.Size = New-Object System.Drawing.Size(245, 22)
$appointmentHint.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 8)
$appointmentHint.ForeColor = $Script:ThemeMutedText
$appointmentCard.Controls.Add($appointmentHint)

$appointmentToolTip = New-Object System.Windows.Forms.ToolTip
$appointmentToolTip.SetToolTip($appointmentButton, 'Creates one appointment file per refill date, with an alert one day before each appointment.')

#endregion

#region Supply calculations and iCalendar text helpers
function Get-RefillDate([datetime]$FillDate, [int]$Quantity, [int]$Doses, [string]$Frequency) {
    if ($Frequency -eq 'Weekly')
    {
        return $FillDate.Date.AddDays(($Quantity - 1) * 7)
    }

    return $FillDate.Date.AddDays([Math]::Ceiling($Quantity / $Doses) - 1)
}

function Get-RemainingQuantity([int]$Quantity, [int]$Doses, [string]$Frequency, [int]$DaysElapsed) {
    if ($Frequency -eq 'Weekly')
    {
        $dosesTaken = if ($DaysElapsed -gt 0) { [Math]::Floor(($DaysElapsed - 1) / 7) + 1 } else { 0 }
        return [Math]::Max(0, $Quantity - $dosesTaken)
    }

    return [Math]::Max(0, $Quantity - ($DaysElapsed * $Doses))
}

function ConvertTo-IcsText([string]$Text) {
    $escapedText = $Text.Replace('\', '\\')
    $escapedText = $escapedText.Replace(';', '\;')
    $escapedText = $escapedText.Replace(',', '\,')
    $escapedText = $escapedText.Replace("`r`n", '\n')
    $escapedText = $escapedText.Replace("`r", '\n')
    $escapedText = $escapedText.Replace("`n", '\n')
    return $escapedText
}

function ConvertTo-IcsContentLine([string]$Line) {
    $builder = New-Object System.Text.StringBuilder
    $lineBytes = 0
    $elements = [System.Globalization.StringInfo]::GetTextElementEnumerator($Line)
    while ($elements.MoveNext())
    {
        $element = $elements.GetTextElement()
        $elementBytes = [System.Text.Encoding]::UTF8.GetByteCount($element)
        if ($lineBytes + $elementBytes -gt 75)
        {
            [void]$builder.Append("`r`n ")
            $lineBytes = 1
        }
        [void]$builder.Append($element)
        $lineBytes += $elementBytes
    }
    return $builder.ToString()
}

#endregion

#region Medication result calculations
function Get-MedicationResults([datetime]$AsOfDate) {
    $results = [System.Collections.Generic.List[System.Object]]::new()
    $count = [int]$medicationCountInput.Value
    for ($index = 0; $index -lt $count; $index++)
    {
        $entry = $Script:MedicationEntries[$index]
        $name = $entry.NameInput.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name))
        {
            $name = "Medication $($index + 1)"
        }

        $fillDate = $entry.FillDatePicker.Value.Date
        $quantity = [int]$entry.QuantityInput.Value
        $doses = [int]$entry.DosesInput.Value
        $frequency = [string]$entry.FrequencyBox.SelectedItem
        $hasStarted = $AsOfDate.Date -ge $fillDate
        $remaining = if ($hasStarted)
        {
            $daysElapsed = ($AsOfDate.Date - $fillDate).Days
            Get-RemainingQuantity -Quantity $quantity -Doses $doses -Frequency $frequency -DaysElapsed $daysElapsed
        }
        else
        {
            0
        }
        $refillDate = Get-RefillDate -FillDate $fillDate -Quantity $quantity -Doses $doses -Frequency $frequency
        [void]$results.Add([PSCustomObject]@{
            Name = $name
            FillDate = $fillDate
            Quantity = $quantity
            Doses = $doses
            Frequency = $frequency
            Remaining = $remaining
            HasStarted = $hasStarted
            RefillDate = $refillDate
        })
    }
    return $results.ToArray()
}

function Update-Result {
    $checkDate = $checkDatePicker.Value.Date
    $results = @(Get-MedicationResults -AsOfDate $checkDate)
    $Script:CurrentMedicationResults = $results
    $resultsHint.Visible = $false
    $resultsList.BeginUpdate()
    try
    {
        $resultsList.Columns[1].Text = 'Qty on ' + $checkDate.ToShortDateString()
        $resultsList.Items.Clear()
        foreach ($medication in $results)
        {
            $availability = if ($medication.HasStarted) { [string]$medication.Remaining } else { 'Not filled yet' }
            $item = New-Object System.Windows.Forms.ListViewItem($medication.Name)
            [void]$item.SubItems.Add($availability)
            [void]$item.SubItems.Add($medication.RefillDate.ToShortDateString())
            $item.ToolTipText = $medication.Name
            [void]$resultsList.Items.Add($item)
        }
    }
    finally
    {
        $resultsList.EndUpdate()
    }
}

#endregion

#region Monthly medication availability calendar
function Get-CalendarDaySummary([datetime]$CalendarDate, [object[]]$Medications) {
    $medicationLines = [System.Collections.Generic.List[System.Object]]::new()
    foreach ($medication in $Medications)
    {
        if ($CalendarDate.Date -lt $medication.FillDate)
        {
            $availability = 'Not filled'
            $status = 'NotFilled'
        }
        else
        {
            $daysElapsed = ($CalendarDate.Date - $medication.FillDate).Days
            $remaining = Get-RemainingQuantity -Quantity $medication.Quantity -Doses $medication.Doses -Frequency $medication.Frequency -DaysElapsed $daysElapsed
            $availability = [string]$remaining
            $status = if ($remaining -gt 0) { 'Active' } else { 'Depleted' }
        }
        [void]$medicationLines.Add([PSCustomObject]@{
            Text = ('{0} | {1}' -f $medication.Name, $availability)
            Status = $status
        })
    }

    return [PSCustomObject]@{
        Lines = $medicationLines.ToArray()
    }
}

function Show-AvailabilityCalendar {
    $checkDate = $checkDatePicker.Value.Date
    $medications = @(Get-MedicationResults -AsOfDate $checkDate)
    $themeMutedText = $Script:ThemeMutedText
    $themeGreenText = $Script:ThemeGreenText
    $themeRedText = $Script:ThemeRedText
    $themeText = $Script:ThemeText
    $themeSurface = $Script:ThemeSurface
    $workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $calendarWidth = [Math]::Min(1280, $workingArea.Width)
    $calendarHeight = [Math]::Min(980, $workingArea.Height)
    $calendarForm = New-Object System.Windows.Forms.Form
    $calendarForm.Text = 'Medication availability'
    $calendarForm.Size = New-Object System.Drawing.Size($calendarWidth, $calendarHeight)
    $calendarForm.StartPosition = 'CenterParent'
    $calendarForm.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 9)
    $calendarForm.BackColor = $Script:ThemeBackground
    $calendarForm.ForeColor = $Script:ThemeText

    $previousButton = New-Object System.Windows.Forms.Button
    $previousButton.Text = '< Previous'
    $previousButton.Location = New-Object System.Drawing.Point(12, 14)
    $previousButton.Size = New-Object System.Drawing.Size(100, 30)
    Set-ButtonStyle -Button $previousButton
    $calendarForm.Controls.Add($previousButton)

    $calendarTitle = New-Object System.Windows.Forms.Label
    $calendarTitle.Location = New-Object System.Drawing.Point(120, 12)
    $calendarTitle.Size = New-Object System.Drawing.Size(($calendarWidth - 244), 34)
    $calendarTitle.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 14, [System.Drawing.FontStyle]::Bold)
    $calendarTitle.ForeColor = $Script:ThemeAccent
    $calendarTitle.TextAlign = 'MiddleCenter'
    $calendarTitle.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $calendarForm.Controls.Add($calendarTitle)

    $nextButton = New-Object System.Windows.Forms.Button
    $nextButton.Text = 'Next >'
    $nextButton.Location = New-Object System.Drawing.Point(($calendarWidth - 120), 14)
    $nextButton.Size = New-Object System.Drawing.Size(100, 30)
    $nextButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    Set-ButtonStyle -Button $nextButton
    $calendarForm.Controls.Add($nextButton)

    $calendarGrid = New-Object System.Windows.Forms.TableLayoutPanel
    $calendarGrid.Location = New-Object System.Drawing.Point(12, 58)
    $calendarGrid.Size = New-Object System.Drawing.Size(
        [Math]::Max(760, $calendarForm.ClientSize.Width - 24),
        [Math]::Max(420, $calendarForm.ClientSize.Height - 170)
    )
    $calendarGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $calendarGrid.ColumnCount = 7
    $calendarGrid.RowCount = 7
    $calendarGrid.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    $calendarGrid.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
    $calendarGrid.BackColor = $Script:ThemeSurfaceAlt
    for ($column = 0; $column -lt 7; $column++)
    {
        $columnStyle = New-Object System.Windows.Forms.ColumnStyle
        $columnStyle.SizeType = [System.Windows.Forms.SizeType]::Percent
        $columnStyle.Width = 100.0 / 7
        [void]$calendarGrid.ColumnStyles.Add($columnStyle)
    }
    for ($row = 0; $row -lt 7; $row++)
    {
        $rowStyle = New-Object System.Windows.Forms.RowStyle
        if ($row -eq 0)
        {
            $rowStyle.SizeType = [System.Windows.Forms.SizeType]::Absolute
            $rowStyle.Height = 30
        }
        else
        {
            $rowStyle.SizeType = [System.Windows.Forms.SizeType]::Percent
            $rowStyle.Height = 100.0 / 6
        }
        [void]$calendarGrid.RowStyles.Add($rowStyle)
    }
    $calendarForm.Controls.Add($calendarGrid)

    $firstDayOfWeek = [int][System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.FirstDayOfWeek
    $weekdayNames = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.AbbreviatedDayNames
    for ($column = 0; $column -lt 7; $column++)
    {
        $weekdayLabel = New-Object System.Windows.Forms.Label
        $weekdayLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
        $weekdayLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $weekdayLabel.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10, [System.Drawing.FontStyle]::Bold)
        $weekdayLabel.BackColor = $Script:ThemePrimary
        $weekdayLabel.ForeColor = [System.Drawing.Color]::White
        $weekdayLabel.Text = $weekdayNames[($firstDayOfWeek + $column) % 7]
        $calendarGrid.Controls.Add($weekdayLabel, $column, 0)
    }

    $legend = New-Object System.Windows.Forms.Label
    $legend.Location = New-Object System.Drawing.Point(12, ($calendarGrid.Bottom + 10))
    $legend.Size = New-Object System.Drawing.Size($calendarGrid.Width, 50)
    $legend.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $legend.ForeColor = $Script:ThemeText
    $legend.Text = "Each day lists Medication | remaining quantity at the start of that date.`r`nGreen: supply available. Red: depleted. Gray: not filled yet. The outlined date is the shared spot-check date."
    $calendarForm.Controls.Add($legend)

    $calendarToolTip = New-Object System.Windows.Forms.ToolTip
    $calendarToolTip.AutoPopDelay = 30000
    $calendarToolTip.InitialDelay = 500
    $calendarToolTip.ReshowDelay = 100

    $calendarForm.Tag = [datetime]::new($checkDate.Year, $checkDate.Month, 1)
    $refreshCalendar = {
        $monthStart = [datetime]$calendarForm.Tag
        $calendarTitle.Text = $monthStart.ToString('MMMM yyyy')
        foreach ($control in @($calendarGrid.Controls))
        {
            if ($calendarGrid.GetRow($control) -gt 0)
            {
                $calendarGrid.Controls.Remove($control)
                $control.Dispose()
            }
        }

        $firstVisibleDate = $monthStart.AddDays(-(([int]$monthStart.DayOfWeek - $firstDayOfWeek + 7) % 7))
        for ($dayOffset = 0; $dayOffset -lt 42; $dayOffset++)
        {
            $calendarDate = $firstVisibleDate.AddDays($dayOffset)
            $daySummary = Get-CalendarDaySummary -CalendarDate $calendarDate -Medications $medications
            $medicationText = ($daySummary.Lines | ForEach-Object { $_.Text }) -join "`r`n"
            $tooltipText = "$($calendarDate.ToLongDateString())`r`n$medicationText"

            $cell = New-Object System.Windows.Forms.Panel
            $cell.Dock = [System.Windows.Forms.DockStyle]::Fill
            $cell.Margin = New-Object System.Windows.Forms.Padding(1)
            $cell.Padding = New-Object System.Windows.Forms.Padding(3)
            $cell.BackColor = $themeSurface
            $calendarToolTip.SetToolTip($cell, $tooltipText)
            $cell.BorderStyle = if ($calendarDate.Date -eq $checkDate) {
                [System.Windows.Forms.BorderStyle]::Fixed3D
            } else {
                [System.Windows.Forms.BorderStyle]::FixedSingle
            }

            for ($lineIndex = $daySummary.Lines.Count - 1; $lineIndex -ge 0; $lineIndex--)
            {
                $line = $daySummary.Lines[$lineIndex]
                $lineLabel = New-Object System.Windows.Forms.Label
                $lineLabel.Dock = [System.Windows.Forms.DockStyle]::Top
                $lineLabel.Height = 18
                $lineLabel.AutoEllipsis = $true
                $lineLabel.Text = $line.Text
                $lineLabel.ForeColor = if ($line.Status -eq 'Active') {
                    $themeGreenText
                } elseif ($line.Status -eq 'Depleted') {
                    $themeRedText
                } else {
                    $themeMutedText
                }
                $calendarToolTip.SetToolTip($lineLabel, $tooltipText)
                $cell.Controls.Add($lineLabel)
            }
            $dayLabel = New-Object System.Windows.Forms.Label
            $dayLabel.Dock = [System.Windows.Forms.DockStyle]::Top
            $dayLabel.Height = 20
            $dayLabel.Text = [string]$calendarDate.Day
            $dayLabel.ForeColor = if ($calendarDate.Month -ne $monthStart.Month -or $calendarDate.Year -ne $monthStart.Year) { $themeMutedText } else { $themeText }
            $calendarToolTip.SetToolTip($dayLabel, $tooltipText)
            $cell.Controls.Add($dayLabel)
            $calendarGrid.Controls.Add($cell, $dayOffset % 7, ([int][Math]::Floor($dayOffset / 7) + 1))
        }
    }.GetNewClosure()

    $previousButton.Add_Click({
        $calendarForm.Tag = ([datetime]$calendarForm.Tag).AddMonths(-1)
        & $refreshCalendar
    }.GetNewClosure())
    $nextButton.Add_Click({
        $calendarForm.Tag = ([datetime]$calendarForm.Tag).AddMonths(1)
        & $refreshCalendar
    }.GetNewClosure())

    & $refreshCalendar
    try
    {
        [void]$calendarForm.ShowDialog($form)
    }
    finally
    {
        $calendarToolTip.Dispose()
        $calendarForm.Dispose()
    }
}

#endregion

#region iCalendar generation and application event handlers
function New-IcsEventContent([object[]]$Medications, [TimeSpan]$AppointmentTime, [int]$DaysBeforeRefill) {
    if ($Medications.Count -eq 0)
    {
        throw 'At least one medication is required for an appointment.'
    }
    $refillDate = $Medications[0].RefillDate.Date
    foreach ($medication in $Medications)
    {
        if ($medication.RefillDate.Date -ne $refillDate)
        {
            throw 'Medications in an appointment must have the same refill date.'
        }
    }
    if ($DaysBeforeRefill -lt 0 -or $DaysBeforeRefill -gt ($refillDate - [datetime]::MinValue.Date).Days)
    {
        throw "The days-before value is outside the valid date range for refill date $($refillDate.ToShortDateString())."
    }
    $appointmentDate = $refillDate.AddDays(-$DaysBeforeRefill)
    $appointmentStartLocal = [DateTime]::SpecifyKind(
        $appointmentDate.Add($AppointmentTime),
        [DateTimeKind]::Local
    )
    if ([System.TimeZoneInfo]::Local.IsInvalidTime($appointmentStartLocal))
    {
        throw "The selected appointment time does not exist on $($appointmentDate.ToShortDateString()) because of a daylight-saving time change."
    }

    $appointmentStartUtc = $appointmentStartLocal.ToUniversalTime()
    $appointmentEndUtc = $appointmentStartUtc.AddHours(1)
    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    $startUtc = $appointmentStartUtc.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $endUtc = $appointmentEndUtc.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $stampUtc = [DateTime]::UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $uid = '{0}@meds-calculator' -f [guid]::NewGuid().ToString()
    $names = ($Medications | ForEach-Object { $_.Name }) -join ', '
    $escapedSummary = ConvertTo-IcsText -Text "$names refill"
    $refillDateText = $refillDate.ToString('yyyy-MM-dd', $invariantCulture)
    $escapedDescription = ConvertTo-IcsText -Text "Refill $names medication supply. Refill date: $refillDateText."
    $eventLines = @(
        'BEGIN:VEVENT'
        "UID:$uid"
        "DTSTAMP:$stampUtc"
        "DTSTART:$startUtc"
        "DTEND:$endUtc"
        "SUMMARY:$escapedSummary"
        "DESCRIPTION:$escapedDescription"
        'BEGIN:VALARM'
        'TRIGGER;RELATED=START:-P1D'
        'ACTION:DISPLAY'
        'DESCRIPTION:Medication refill reminder'
        'END:VALARM'
        'END:VEVENT'
    )
    return ($eventLines | ForEach-Object { ConvertTo-IcsContentLine -Line $_ }) -join "`r`n"
}

function New-IcsCalendarContent([object[]]$Medications, [TimeSpan]$AppointmentTime, [int]$DaysBeforeRefill) {
    $calendarLines = [System.Collections.Generic.List[string]]::new()
    foreach ($line in @(
        'BEGIN:VCALENDAR'
        'VERSION:2.0'
        'PRODID:-//Meds Calculator//Refill Appointment//EN'
        'CALSCALE:GREGORIAN'
        'METHOD:PUBLISH'
    ))
    {
        [void]$calendarLines.Add($line)
    }
    [void]$calendarLines.Add((New-IcsEventContent -Medications $Medications -AppointmentTime $AppointmentTime -DaysBeforeRefill $DaysBeforeRefill))
    [void]$calendarLines.Add('END:VCALENDAR')
    return $calendarLines -join "`r`n"
}

function New-IcsAppointmentFiles([object[]]$Medications, [TimeSpan]$AppointmentTime, [int]$DaysBeforeRefill) {
    $groups = @($Medications | Group-Object -Property { $_.RefillDate.Date.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) } | Sort-Object Name)
    foreach ($group in $groups)
    {
        [PSCustomObject]@{
            FileName = "Refills-$($group.Name).ics"
            Content = New-IcsCalendarContent -Medications $group.Group -AppointmentTime $AppointmentTime -DaysBeforeRefill $DaysBeforeRefill
        }
    }
}

$copyFillDateButton.Add_Click({
    $fillDate = $Script:MedicationEntries[0].FillDatePicker.Value.Date
    for ($index = 1; $index -lt [int]$medicationCountInput.Value; $index++)
    {
        $Script:MedicationEntries[$index].FillDatePicker.Value = $fillDate
    }
    if ($Script:CurrentMedicationResults.Count -gt 0)
    {
        Update-Result
    }
})
$calculateButton.Add_Click({ Update-Result })
$calendarButton.Add_Click({ Show-AvailabilityCalendar })
$appointmentButton.Add_Click({
    Update-Result
    $medications = @(Get-MedicationResults -AsOfDate $checkDatePicker.Value.Date)
    try
    {
        $calendarFiles = @(New-IcsAppointmentFiles -Medications $medications -AppointmentTime $appointmentTimePicker.Value.TimeOfDay -DaysBeforeRefill ([int]$daysBeforeInput.Value))
    }
    catch
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            "Could not create the appointments: $($_.Exception.Message)",
            'Appointment creation failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Choose a folder for the medication refill appointment files.'
    $documentsPath = [Environment]::GetFolderPath('MyDocuments')
    if (-not [string]::IsNullOrWhiteSpace($documentsPath))
    {
        $dialog.SelectedPath = $documentsPath
    }

    try
    {
        if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK)
        {
            return
        }
        $destinationFolder = $dialog.SelectedPath
    }
    finally
    {
        $dialog.Dispose()
    }

    $savedPaths = [System.Collections.Generic.List[string]]::new()
    $currentPath = $null
    try
    {
        foreach ($calendarFile in $calendarFiles)
        {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($calendarFile.FileName)
            $counter = 1
            do
            {
                $suffix = if ($counter -gt 1) { "-$counter" } else { '' }
                $currentPath = [System.IO.Path]::Combine($destinationFolder, "$baseName$suffix.ics")
                $counter++
            } while (Test-Path -LiteralPath $currentPath)

            $stream = [System.IO.File]::Open($currentPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try
            {
                $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
                try
                {
                    $writer.Write($calendarFile.Content + "`r`n")
                }
                finally
                {
                    $writer.Dispose()
                }
            }
            finally
            {
                $stream.Dispose()
            }
            [void]$savedPaths.Add($currentPath)
        }
    }
    catch
    {
        $savedDetails = if ($savedPaths.Count -gt 0) { $savedPaths -join "`r`n" } else { '(none)' }
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            "Could not create all appointment files: $($_.Exception.Message)`r`n`r`nLast attempted file (which may be incomplete):`r`n$currentPath`r`n`r`nFiles created so far:`r`n$savedDetails",
            'Appointment files failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $openErrors = [System.Collections.Generic.List[string]]::new()
    foreach ($savedPath in $savedPaths)
    {
        try
        {
            Start-Process -FilePath $savedPath -ErrorAction Stop
        }
        catch
        {
            [void]$openErrors.Add("$savedPath`r`n$($_.Exception.Message)")
        }
    }
    if ($openErrors.Count -gt 0)
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            "The files were saved, but these could not be opened automatically:`r`n`r`n$($openErrors -join "`r`n`r`n")",
            'Appointment files saved',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})

[void]$form.ShowDialog()
$form.Dispose()
$Script:DatePickerFont.Dispose()
if ($null -ne $applicationIcon)
{
    $applicationIcon.Dispose()
}
#endregion

