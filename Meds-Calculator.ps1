<#
.SYNOPSIS
Estimates medication supply and creates refill calendar appointments.

.DESCRIPTION
Medication Supply Calculator is a Windows Forms application for tracking up to
10 medications. It estimates remaining quantities and refill dates from each
medication's fill date, quantity, and daily or weekly schedule. It can also
show a monthly availability overview and create one Outlook-compatible .ics
file containing a one-hour, one-day-reminder event for each medication.

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
$Script:ThemePrimary = [System.Drawing.Color]::FromArgb(0, 101, 108)
$Script:ThemePrimaryHover = [System.Drawing.Color]::FromArgb(0, 81, 88)
$Script:ThemeAccent = [System.Drawing.Color]::FromArgb(0, 108, 115)
$Script:ThemeText = [System.Drawing.Color]::FromArgb(43, 52, 59)
$Script:ThemeMutedText = [System.Drawing.Color]::FromArgb(91, 105, 112)
$Script:ThemeBackground = [System.Drawing.Color]::FromArgb(238, 242, 244)
$Script:ThemeSurface = [System.Drawing.Color]::FromArgb(255, 255, 255)
$Script:ThemeSurfaceAlt = [System.Drawing.Color]::FromArgb(247, 249, 250)
$Script:ThemeInput = [System.Drawing.Color]::FromArgb(250, 251, 252)
$Script:ThemeBorder = [System.Drawing.Color]::FromArgb(192, 203, 208)
$Script:ThemeHoverSurface = [System.Drawing.Color]::FromArgb(231, 241, 242)
$Script:ThemeNeutral = [System.Drawing.Color]::FromArgb(229, 234, 237)
$Script:ThemeGreen = [System.Drawing.Color]::FromArgb(226, 243, 232)
$Script:ThemeGreenText = [System.Drawing.Color]::FromArgb(37, 105, 66)
$Script:ThemeRed = [System.Drawing.Color]::FromArgb(252, 231, 229)
$Script:ThemeRedText = [System.Drawing.Color]::FromArgb(151, 59, 54)
$availableFontNames = @([System.Drawing.FontFamily]::Families | ForEach-Object { $_.Name })
$Script:UiFontFamily = if ($availableFontNames -contains 'Segoe UI Variable Text') {
    'Segoe UI Variable Text'
} elseif ($availableFontNames -contains 'Segoe UI Variable') {
    'Segoe UI Variable'
} elseif ($availableFontNames -contains 'Aptos') {
    'Aptos'
} elseif ($availableFontNames -contains 'Roboto') {
    'Roboto'
} else {
    'Segoe UI'
}
$Script:DatePickerFont = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Medication Calculator'
$form.Size = New-Object System.Drawing.Size(760, 740)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10)
$form.BackColor = $Script:ThemeBackground
$form.ForeColor = $Script:ThemeText

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
    $heading.ForeColor = $Script:ThemeAccent
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

function Set-ButtonStyle([System.Windows.Forms.Button]$Button, [bool]$Primary = $false) {
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatAppearance.BorderSize = if ($Primary) { 0 } else { 1 }
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 10, [System.Drawing.FontStyle]::Bold)
    if ($Primary)
    {
        $Button.BackColor = $Script:ThemePrimary
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatAppearance.MouseOverBackColor = $Script:ThemePrimaryHover
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
$headerPanel.Size = New-Object System.Drawing.Size(744, 70)
$headerPanel.BackColor = $Script:ThemePrimary
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
$instructionLabel.ForeColor = [System.Drawing.Color]::FromArgb(205, 235, 235)
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
$countHint.Text = "Choose 1-$($Script:MaxMedicationCount); entries below are saved when hidden."
$countHint.Location = New-Object System.Drawing.Point(265, 11)
$countHint.Size = New-Object System.Drawing.Size(430, 22)
$countHint.ForeColor = $Script:ThemeMutedText
$countPanel.Controls.Add($countHint)

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

function New-MedicationEntry([int]$Number) {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(680, 100)
    $panel.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 5)
    $panel.BackColor = $Script:ThemeSurface
    $panel.ForeColor = $Script:ThemeText
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $heading = New-Object System.Windows.Forms.Label
    $heading.Text = "Medication $Number"
    $heading.Location = New-Object System.Drawing.Point(10, 5)
    $heading.Size = New-Object System.Drawing.Size(650, 20)
    $heading.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 9, [System.Drawing.FontStyle]::Bold)
    $heading.ForeColor = $Script:ThemeAccent
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

    Add-FieldLabel -Parent $panel -Text 'Doses taken:' -X 240 -Y 69 -Width 75
    $dosesInput = New-Object System.Windows.Forms.NumericUpDown
    $dosesInput.Minimum = 1
    $dosesInput.Maximum = 4
    $dosesInput.Value = 1
    $dosesInput.Location = New-Object System.Drawing.Point(315, 67)
    $dosesInput.Size = New-Object System.Drawing.Size(65, 25)
    $panel.Controls.Add($dosesInput)

    $frequencyBox = New-Object System.Windows.Forms.ComboBox
    $frequencyBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$frequencyBox.Items.AddRange([object[]]@('Daily', 'Weekly'))
    $frequencyBox.SelectedIndex = 0
    $frequencyBox.Location = New-Object System.Drawing.Point(390, 67)
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
$sharedCard = New-SectionCard -Title 'Shared date and appointment time' -X 16 -Y 370 -Width 712 -Height 74
Add-FieldLabel -Parent $sharedCard -Text 'Spot check date:' -X 12 -Y 36 -Width 110
$checkDatePicker = New-Object System.Windows.Forms.DateTimePicker
$checkDatePicker.Format = 'Short'
$checkDatePicker.Location = New-Object System.Drawing.Point(128, 33)
$checkDatePicker.AutoSize = $false
$checkDatePicker.Size = New-Object System.Drawing.Size(180, 28)
$checkDatePicker.Font = $Script:DatePickerFont
$checkDatePicker.CalendarMonthBackground = $Script:ThemeInput
$checkDatePicker.CalendarForeColor = $Script:ThemeText
$sharedCard.Controls.Add($checkDatePicker)

Add-FieldLabel -Parent $sharedCard -Text 'Appointment time:' -X 320 -Y 36 -Width 125
$appointmentTimePicker = New-Object System.Windows.Forms.DateTimePicker
$appointmentTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$appointmentTimePicker.ShowUpDown = $true
$appointmentTimePicker.Value = (Get-Date).Date.AddHours(9)
$appointmentTimePicker.Location = New-Object System.Drawing.Point(450, 33)
$appointmentTimePicker.AutoSize = $false
$appointmentTimePicker.Size = New-Object System.Drawing.Size(140, 28)
$appointmentTimePicker.BackColor = $Script:ThemeInput
$appointmentTimePicker.ForeColor = $Script:ThemeText
$appointmentTimePicker.CalendarMonthBackground = $Script:ThemeInput
$appointmentTimePicker.CalendarForeColor = $Script:ThemeText
$sharedCard.Controls.Add($appointmentTimePicker)

$sharedHint = New-Object System.Windows.Forms.Label
$sharedHint.Text = '1 hr; 1-day reminder'
$sharedHint.Location = New-Object System.Drawing.Point(598, 36)
$sharedHint.Size = New-Object System.Drawing.Size(102, 22)
$sharedHint.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 8)
$sharedHint.ForeColor = $Script:ThemeMutedText
$sharedCard.Controls.Add($sharedHint)

$calculateButton = New-Object System.Windows.Forms.Button
$calculateButton.Text = 'Calculate all medications'
$calculateButton.Location = New-Object System.Drawing.Point(16, 454)
$calculateButton.Size = New-Object System.Drawing.Size(712, 38)
$calculateButton.TabIndex = 1
Set-ButtonStyle -Button $calculateButton -Primary $true
$form.Controls.Add($calculateButton)
$form.AcceptButton = $calculateButton

$resultPanel = New-SectionCard -Title 'Remaining quantity and refill date' -X 16 -Y 502 -Width 712 -Height 110
$resultsTextBox = New-Object System.Windows.Forms.TextBox
$resultsTextBox.Multiline = $true
$resultsTextBox.ReadOnly = $true
$resultsTextBox.WordWrap = $false
$resultsTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$resultsTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$resultsTextBox.BackColor = $Script:ThemeInput
$resultsTextBox.ForeColor = $Script:ThemeText
$resultsTextBox.Font = New-Object System.Drawing.Font($Script:UiFontFamily, 9)
$resultsTextBox.Location = New-Object System.Drawing.Point(14, 31)
$resultsTextBox.Size = New-Object System.Drawing.Size(680, 68)
$resultsTextBox.Text = 'Enter medication details, then calculate to see each remaining quantity.'
$resultPanel.Controls.Add($resultsTextBox)

$appointmentButton = New-Object System.Windows.Forms.Button
$appointmentButton.Text = 'Create refill appointments (.ics)'
$appointmentButton.Location = New-Object System.Drawing.Point(16, 622)
$appointmentButton.Size = New-Object System.Drawing.Size(344, 36)
$appointmentButton.TabIndex = 2
Set-ButtonStyle -Button $appointmentButton
$form.Controls.Add($appointmentButton)

$calendarButton = New-Object System.Windows.Forms.Button
$calendarButton.Text = 'View availability calendar'
$calendarButton.Location = New-Object System.Drawing.Point(384, 622)
$calendarButton.Size = New-Object System.Drawing.Size(344, 36)
$calendarButton.TabIndex = 3
Set-ButtonStyle -Button $calendarButton
$form.Controls.Add($calendarButton)

$appointmentToolTip = New-Object System.Windows.Forms.ToolTip
$appointmentToolTip.SetToolTip($appointmentButton, 'Creates one one-hour Outlook appointment per medication, with a reminder one day before.')

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

function Get-SafeFileName([string]$Name) {
    foreach ($invalidCharacter in [System.IO.Path]::GetInvalidFileNameChars())
    {
        $Name = $Name.Replace([string]$invalidCharacter, '_')
    }

    $safeName = $Name.Trim().TrimEnd([char[]]@(' ', '.'))
    if ([string]::IsNullOrWhiteSpace($safeName))
    {
        return 'Medication'
    }
    return $safeName
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
    $resultsTextBox.Text = @(
        foreach ($medication in $results)
        {
            $availability = if ($medication.HasStarted) { "$($medication.Remaining) remaining" } else { 'not filled yet' }
            '{0} | {1} | Refill: {2}' -f $medication.Name, $availability, $medication.RefillDate.ToShortDateString()
        }
    ) -join [Environment]::NewLine
}

#endregion

#region Monthly medication availability calendar
function Get-CalendarDaySummary([datetime]$CalendarDate, [object[]]$Medications) {
    $medicationLines = [System.Collections.Generic.List[string]]::new()
    $filledCount = 0
    $activeCount = 0
    foreach ($medication in $Medications)
    {
        if ($CalendarDate.Date -lt $medication.FillDate)
        {
            $availability = 'Not filled'
        }
        else
        {
            $filledCount++
            $daysElapsed = ($CalendarDate.Date - $medication.FillDate).Days
            $remaining = Get-RemainingQuantity -Quantity $medication.Quantity -Doses $medication.Doses -Frequency $medication.Frequency -DaysElapsed $daysElapsed
            if ($remaining -gt 0) { $activeCount++ }
            $availability = [string]$remaining
        }
        [void]$medicationLines.Add(('{0} | {1}' -f $medication.Name, $availability))
    }

    $status = if ($filledCount -eq 0) {
        'NotFilled'
    } elseif ($activeCount -gt 0) {
        'Active'
    } else {
        'Depleted'
    }
    return [PSCustomObject]@{
        Lines = $medicationLines.ToArray()
        FilledCount = $filledCount
        ActiveCount = $activeCount
        Status = $status
    }
}

function Show-AvailabilityCalendar {
    $checkDate = $checkDatePicker.Value.Date
    $medications = @(Get-MedicationResults -AsOfDate $checkDate)
    $uiFontFamily = $Script:UiFontFamily
    $themeMutedText = $Script:ThemeMutedText
    $themeNeutral = $Script:ThemeNeutral
    $themeGreen = $Script:ThemeGreen
    $themeGreenText = $Script:ThemeGreenText
    $themeRed = $Script:ThemeRed
    $themeRedText = $Script:ThemeRedText
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
    $legend.Text = "Each day lists Medication | remaining quantity (or Not filled) at the start of that date.`r`nGreen: at least one supply is active. Red: all filled supplies are empty. Gray: before the first fill date. The outlined date is the shared spot-check date."
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
            $medicationText = $daySummary.Lines -join "`r`n"

            $cell = New-Object System.Windows.Forms.Label
            $cell.Dock = [System.Windows.Forms.DockStyle]::Fill
            $cell.Margin = New-Object System.Windows.Forms.Padding(1)
            $cell.Padding = New-Object System.Windows.Forms.Padding(5)
            $cell.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
            $cell.Font = New-Object System.Drawing.Font($uiFontFamily, 9)
            $cell.AutoEllipsis = $true
            $cell.Text = "$($calendarDate.Day)`r`n$medicationText"
            $calendarToolTip.SetToolTip($cell, "$($calendarDate.ToLongDateString())`r`n$medicationText")
            $cell.BorderStyle = if ($calendarDate.Date -eq $checkDate) {
                [System.Windows.Forms.BorderStyle]::Fixed3D
            } else {
                [System.Windows.Forms.BorderStyle]::FixedSingle
            }

            if ($daySummary.Status -eq 'NotFilled')
            {
                $cell.BackColor = $themeNeutral
                $cell.ForeColor = $themeMutedText
            }
            elseif ($daySummary.Status -eq 'Active')
            {
                $cell.BackColor = $themeGreen
                $cell.ForeColor = $themeGreenText
            }
            else
            {
                $cell.BackColor = $themeRed
                $cell.ForeColor = $themeRedText
            }

            if ($calendarDate.Month -ne $monthStart.Month -or $calendarDate.Year -ne $monthStart.Year)
            {
                $cell.ForeColor = $themeMutedText
            }
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
function New-IcsEventContent($Medication, [TimeSpan]$AppointmentTime) {
    $appointmentStartLocal = [DateTime]::SpecifyKind(
        $Medication.RefillDate.Date.Add($AppointmentTime),
        [DateTimeKind]::Local
    )
    if ([System.TimeZoneInfo]::Local.IsInvalidTime($appointmentStartLocal))
    {
        throw "The selected appointment time does not exist on $($Medication.RefillDate.ToShortDateString()) because of a daylight-saving time change."
    }

    $appointmentStartUtc = $appointmentStartLocal.ToUniversalTime()
    $appointmentEndUtc = $appointmentStartUtc.AddHours(1)
    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    $startUtc = $appointmentStartUtc.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $endUtc = $appointmentEndUtc.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $stampUtc = [DateTime]::UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", $invariantCulture)
    $uid = '{0}@meds-calculator' -f [guid]::NewGuid().ToString()
    $escapedSummary = ConvertTo-IcsText -Text "$($Medication.Name) refill"
    $escapedDescription = ConvertTo-IcsText -Text "Refill $($Medication.Name) medication supply."
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

function New-IcsCalendarContent([object[]]$Medications, [TimeSpan]$AppointmentTime) {
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
    foreach ($medication in $Medications)
    {
        [void]$calendarLines.Add((New-IcsEventContent -Medication $medication -AppointmentTime $AppointmentTime))
    }
    [void]$calendarLines.Add('END:VCALENDAR')
    return $calendarLines -join "`r`n"
}

$calculateButton.Add_Click({ Update-Result })
$calendarButton.Add_Click({ Show-AvailabilityCalendar })
$appointmentButton.Add_Click({
    Update-Result
    $medications = @(Get-MedicationResults -AsOfDate $checkDatePicker.Value.Date)
    try
    {
        $calendarContent = New-IcsCalendarContent -Medications $medications -AppointmentTime $appointmentTimePicker.Value.TimeOfDay
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

    $safeNames = @($medications | ForEach-Object { Get-SafeFileName -Name $_.Name })
    $fileNameBase = Get-SafeFileName -Name ($safeNames -join '-')
    if ($fileNameBase.Length -gt 120)
    {
        $fileNameBase = $fileNameBase.Substring(0, 120).TrimEnd([char[]]@(' ', '.'))
    }
    if ($medications.Count -eq 1)
    {
        $suggestedFileName = "$fileNameBase-Refill-$($medications[0].RefillDate.ToString('yyyy-MM-dd')).ics"
    }
    else
    {
        $suggestedFileName = "$fileNameBase-Refills.ics"
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Save medication refill appointments'
    $dialog.Filter = 'iCalendar file (*.ics)|*.ics'
    $dialog.DefaultExt = 'ics'
    $dialog.AddExtension = $true
    $dialog.FileName = $suggestedFileName
    $documentsPath = [Environment]::GetFolderPath('MyDocuments')
    if (-not [string]::IsNullOrWhiteSpace($documentsPath))
    {
        $dialog.InitialDirectory = $documentsPath
    }

    $savedFilePath = $null
    try
    {
        if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK)
        {
            return
        }

        $savedFilePath = $dialog.FileName
        [System.IO.File]::WriteAllText(
            $savedFilePath,
            $calendarContent + "`r`n",
            [System.Text.UTF8Encoding]::new($false)
        )
    }
    catch
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            "Could not create the appointment file: $($_.Exception.Message)",
            'Appointment file failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    finally
    {
        $dialog.Dispose()
    }

    try
    {
        Start-Process -FilePath $savedFilePath -ErrorAction Stop
    }
    catch
    {
        [void][System.Windows.Forms.MessageBox]::Show(
            $form,
            "The appointment file was saved, but could not be opened automatically:`r`n$savedFilePath`r`n`r`n$($_.Exception.Message)",
            'Appointment file saved',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})

[void]$form.ShowDialog()
$form.Dispose()
$Script:DatePickerFont.Dispose()
#endregion

