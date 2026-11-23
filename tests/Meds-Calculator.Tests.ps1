$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Meds-Calculator.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0)
{
    throw "Could not parse calculator script: $($parseErrors -join '; ')"
}

$functionNames = @(
    'Get-RemainingQuantity'
    'ConvertTo-IcsText'
    'ConvertTo-IcsContentLine'
    'Get-CalendarDaySummary'
    'New-IcsEventContent'
    'New-IcsCalendarContent'
    'New-IcsAppointmentFiles'
)
foreach ($functionName in $functionNames)
{
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    }, $true)
    if ($null -eq $definition)
    {
        throw "Missing function: $functionName"
    }
    . ([scriptblock]::Create($definition.Extent.Text))
}

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition)
    {
        throw $Message
    }
}

$alpha = [PSCustomObject]@{
    Name = 'Alpha'; FillDate = [datetime]'2026-01-01'; Quantity = 2
    Doses = 1; Frequency = 'Daily'; RefillDate = [datetime]'2026-01-02'
}
$beta = [PSCustomObject]@{
    Name = 'Beta; Plus'; FillDate = [datetime]'2026-01-01'; Quantity = 2
    Doses = 1; Frequency = 'Daily'; RefillDate = [datetime]'2026-01-02'
}
$gamma = [PSCustomObject]@{
    Name = 'Gamma'; FillDate = [datetime]'2026-01-01'; Quantity = 5
    Doses = 1; Frequency = 'Daily'; RefillDate = [datetime]'2026-01-05'
}
$future = [PSCustomObject]@{
    Name = 'Future'; FillDate = [datetime]'2026-01-10'; Quantity = 3
    Doses = 1; Frequency = 'Daily'; RefillDate = [datetime]'2026-01-12'
}

$day = Get-CalendarDaySummary -CalendarDate ([datetime]'2026-01-03') -Medications @($alpha, $gamma, $future)
Assert-True ($day.Lines.Count -eq 3) 'Expected one calendar line per medication.'
Assert-True ($day.Lines[0].Text -eq 'Alpha | 0' -and $day.Lines[0].Status -eq 'Depleted') 'Depleted medication status is incorrect.'
Assert-True ($day.Lines[1].Text -eq 'Gamma | 3' -and $day.Lines[1].Status -eq 'Active') 'Active medication status is incorrect.'
Assert-True ($day.Lines[2].Text -eq 'Future | Not filled' -and $day.Lines[2].Status -eq 'NotFilled') 'Not-filled medication status is incorrect.'

$time = [TimeSpan]::FromHours(9)
$files = @(New-IcsAppointmentFiles -Medications @($alpha, $beta, $gamma) -AppointmentTime $time -DaysBeforeRefill 1)
Assert-True ($files.Count -eq 2) 'Expected one file for each distinct refill date.'
Assert-True ($files[0].FileName -eq 'Refills-2026-01-02.ics' -and $files[1].FileName -eq 'Refills-2026-01-05.ics') 'Refill-date filenames are incorrect.'
$first = $files[0].Content.Replace("`r`n ", '')
$second = $files[1].Content.Replace("`r`n ", '')
Assert-True ([regex]::Matches($first, 'BEGIN:VEVENT').Count -eq 1) 'Shared-date medications must have one event.'
Assert-True ($first.Contains('SUMMARY:Alpha\, Beta\; Plus refill')) 'Shared-date event title or escaping is incorrect.'
Assert-True ($first.Contains('Refill date: 2026-01-02.') -and $first.Contains('TRIGGER;RELATED=START:-P1D')) 'Refill date or one-day alert is missing.'
Assert-True ($second.Contains('SUMMARY:Gamma refill')) 'Separate-date event title is incorrect.'

$culture = [System.Globalization.CultureInfo]::InvariantCulture
$expectedStart = 'DTSTART:' + [DateTime]::SpecifyKind(([datetime]'2026-01-01').Add($time), [DateTimeKind]::Local).ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", $culture)
Assert-True ($first.Contains($expectedStart)) 'The days-before offset was not applied to the appointment date.'
$refillDay = New-IcsCalendarContent -Medications @($alpha) -AppointmentTime $time -DaysBeforeRefill 0
$expectedRefillStart = 'DTSTART:' + [DateTime]::SpecifyKind(([datetime]'2026-01-02').Add($time), [DateTimeKind]::Local).ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", $culture)
Assert-True ($refillDay.Contains($expectedRefillStart)) 'Zero days before should schedule on the refill date.'
$yearBoundary = New-IcsCalendarContent -Medications @($alpha) -AppointmentTime $time -DaysBeforeRefill 2
$expectedPreviousYear = 'DTSTART:' + [DateTime]::SpecifyKind(([datetime]'2025-12-31').Add($time), [DateTimeKind]::Local).ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", $culture)
Assert-True ($yearBoundary.Contains($expectedPreviousYear)) 'Days before should work across a year boundary.'

$longName = 'Long' * 25
$longMedication = [PSCustomObject]@{ Name = $longName; RefillDate = [datetime]'2026-01-02' }
$longEvent = New-IcsCalendarContent -Medications @($alpha, $longMedication) -AppointmentTime $time -DaysBeforeRefill 0
Assert-True ($longEvent.Contains("`r`n ")) 'Long event lines should be folded.'
Assert-True ($longEvent.Replace("`r`n ", '').Contains("SUMMARY:Alpha\, $longName refill")) 'Folded event title lost medication names.'
foreach ($line in ($longEvent -split "`r`n"))
{
    Assert-True ([System.Text.Encoding]::UTF8.GetByteCount($line) -le 75) 'An iCalendar content line exceeds 75 bytes.'
}

$rejected = $false
try
{
    New-IcsCalendarContent -Medications @($alpha, $gamma) -AppointmentTime $time -DaysBeforeRefill 0 | Out-Null
}
catch
{
    if ($_.Exception.Message -notlike '*same refill date*') { throw }
    $rejected = $true
}
Assert-True $rejected 'An appointment must reject medications with different refill dates.'

foreach ($invalidOffset in @(-1, 800000))
{
    $rejected = $false
    try
    {
        New-IcsCalendarContent -Medications @($alpha) -AppointmentTime $time -DaysBeforeRefill $invalidOffset | Out-Null
    }
    catch
    {
        if ($_.Exception.Message -notlike '*outside the valid date range*') { throw }
        $rejected = $true
    }
    Assert-True $rejected 'An invalid days-before offset must be rejected.'
}

Write-Output 'PASS: calendar statuses and grouped iCalendar appointments'
