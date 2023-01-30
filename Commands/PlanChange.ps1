<# This command (custom cmdlet) is used for upgrading or downgrading the plan of the specified mobile number by automatically changing Column 5 & 6 on the Ooredoo Master File based on the specified parameters. This will also automatically change the value of the Column 1 (Date Requested/Date Last Modified) to the date it was modified, while the Column 13 will automatically generate new details for the specified action (upgrade or downgrade). #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n" -ForegroundColor DarkRed

Write-Host "Warning: Please enter the exact mobile number that needs to be upgraded or downgraded. If you enter an invalid value, this command will keep running and prompt you for the correct mobile number until it matches a record in the database entries.`n" -ForegroundColor Cyan

# mobile number initialization
[string]$Number = Read-Host "Enter the Mobile Number to Upgrade or Downgrade"

<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "$(Get-Location)\Commands\Database\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)
$LastUsedRow = $MainSheet.UsedRange.Rows.Count

# query (main)
$QueryNumber = $MainSheet.Range("D2:D$LastUsedRow").Find($Number)

# mobile number validation loop
if ($Number.Length -eq 8) {
  while ($QueryNumber.Value2 -ne $Number) {
    [string][ValidateLength(8, 8)][ValidateNotNullOrEmpty()]$Number = Read-Host "Enter the Mobile Number Again"
    # modified query (if validation runs)
    $QueryNumber = $MainSheet.Range("D2:D$LastUsedRow").Find($Number)
  }
}
else {
  while ($QueryNumber.Value2 -ne $Number) {
    [string][ValidateLength(8, 8)][ValidateNotNullOrEmpty()]$Number = Read-Host "Enter the Mobile Number Again"
    # modified query (if validation runs)
    $QueryNumber = $MainSheet.Range("D2:D$LastUsedRow").Find($Number)
  }
}

# row indentifier
$RowIndex = $QueryNumber.Row

# all columns needed
$Col1 = $MainSheet.Cells($RowIndex, 1)   # Date Requested/Date Last Modified
$Col4 = $MainSheet.Cells($RowIndex, 4)   # MSISDN (Mobile Number)
$Col5 = $MainSheet.Cells($RowIndex, 5)   # Current Plan
$Col6 = $MainSheet.Cells($RowIndex, 6)   # Plan Rate
$Col9 = $MainSheet.Cells($RowIndex, 9)   # Name
$Col11 = $MainSheet.Cells($RowIndex, 11) # Staff Grade
$Col12 = $MainSheet.Cells($RowIndex, 12) # Request Completion Date
$Col13 = $MainSheet.Cells($RowIndex, 13) # Remarks

# date and time details used for remarks
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# default value of Column 5: Ooredoo Plan
$PlanDefaultValue = $Col5.Value2

# save the changes Upgrade Function has done
function ProceedUpgrade {

  # this will record the history of upgrade changes
  if ([string]::IsNullorEmpty($Col13.Value2)) {
    $Col13.Value = "$($CurrentDateTime) - From Plan $($PlanDefaultValue) Upgraded to Plan $($Col5.Value2)"
  }
  else {
    $Column13Value = $Col13.Value2
    $Col13.Value = "$($Column13Value)`n$($CurrentDateTime) - From Plan $($PlanDefaultValue) Upgraded to Plan $($Col5.Value2)"
  }

  $Workbook.Save()  # saves the file
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # completed process prompt message
  $Message = "Successfully Upgraded."
  Write-Host $Message -ForegroundColor Green
}

# save the changes Downgrade Function has done
function ProceedDowngrade {
  # this will record the history of downgrade changes
  if ([string]::IsNullorEmpty($Col13.Value2)) {
    $Col13.Value = "$($CurrentDateTime) - From Plan $($PlanDefaultValue) Downgraded to Plan $($Col5.Value2)"
  }
  else {
    $Column13Value = $Col13.Value2
    $Col13.Value = "$($Column13Value)`n$($CurrentDateTime) - From Plan $($PlanDefaultValue) Downgraded to Plan $($Col5.Value2)"
  }

  $Workbook.Save()  # saves the file
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # completed process prompt message
  $Message = "Successfully Downgraded."
  Write-Host $Message -ForegroundColor DarkGreen
}

# cancel function
function Cancel {
  # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
  $Excel.DisplayAlerts = $false
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process
}

# upgrade function
function Upgrade {

  param (
    # upgrade plan
    [Parameter(Mandatory = $true)]
    $UpgradeToPlan
  )

  # date requested - automatic invocation
  $CurrentDate = Get-Date -Format "dd-MMM-yyyy"
  $Col1.Value = $CurrentDate

  # plan upgradation logic
  switch ($UpgradeToPlan) {
    "B" { $Col5.Value = "B"; $Col6.Value = "90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "90" }
    "D" { $Col5.Value = "D"; $Col6.Value = "110.50" }
    "E" { $Col5.Value = "E"; $Col6.Value = "130" }
    "F" { $Col5.Value = "F"; $Col6.Value = "135" }
    "G" { $Col5.Value = "G"; $Col6.Value = "195" }
    "H" { $Col5.Value = "H"; $Col6.Value = "360" }

    Default { $Col5.Value = $PlanDefaultValue }  # nothing to do; remains the same
  }

  # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
  $Col12.Interior.ColorIndex = 6
  for ($i = 1; $i -lt $LastUsedRow; $i++) {
    $Col12Value = "R-$($i)"
    if ($Mainsheet.Range("L2:L$($LastUsedRow)").Value2 -notcontains $Col12Value) {
      $Col12.Value = $Col12Value
      break
    }
    else {
      $Col12.Value = ""
    }
  }
}

# downgrade function
function Downgrade {
  param (
    # downgrade plan
    [Parameter(Mandatory = $true)]
    $DowngradeToPlan
  )

  # date requested - automatic invocation
  $CurrentDate = Get-Date -Format "dd-MMM-yyyy"
  $Col1.Value = $CurrentDate

  # plan downgradation logic
  switch ($DowngradeToPlan) {
    "A" { $Col5.Value = "A"; $Col6.Value = "58.50" }
    "B" { $Col5.Value = "B"; $Col6.Value = "90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "90" }
    "D" { $Col5.Value = "D"; $Col6.Value = "110.50" }
    "E" { $Col5.Value = "E"; $Col6.Value = "130" }
    "F" { $Col5.Value = "F"; $Col6.Value = "135" }
    "G" { $Col5.Value = "G"; $Col6.Value = "195" }

    Default { $Col5.Value = $PlanDefaultValue }  # nothing to do; remains the same
  }

  # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
  $Col12.Interior.ColorIndex = 6
}

# function that displays the necessary information of card holder
function InfoDisplay {
  function EligiblePlan {
    # plan eligibility values
    $NA = "Not Applicable as per the ACIFM Staff Grades and Benefits Section"
    $AC = "Plan A to Plan C only"
    $AD = "Plan A to Plan D only"
    $AG = "Plan A to Plan G only"
    $AH = "Plan A to Plan H"

    # switch statement conditions
    switch ($Col11.Value2) {
      "S1" { $NA }
      "S2" { $NA }
      "T1" { $NA }
      "T2" { $NA }
      "T3" { $NA }

      "S3" { $AC }

      "S4" { $AD }
      "T4A" { $AD }

      "T4B" { $AG }
      "T4C" { $AG }
      "M1A" { $AG }

      "M1B" { $AH }
      "M1C" { $AH }
      "M2A" { $AH }
      "M2B" { $AH }
      "M3" { $AH }
      "M4" { $AH }

      Default { "I don't know about that because the Grade Value is ' $($Col11.Value2) '." }
    }
    return
  }

  # display output format
  Write-Host "`n::::: MASTER FILE DETAILS :::::`n
Sim Holder:              $($Col9.Value2)
Mobile Number:           $($Col4.Value2)
Current Ooredoo Plan:    $PlanDefaultValue
Employee Grade:          $($Col11.Value2)
Plan Eligibility:        $(EligiblePlan)`n" -ForegroundColor Magenta
}

# run InfoDisplay Function first
InfoDisplay

# initial action prompt used for action logic
$Action = Read-Host "Enter 'U' to Upgrade - Enter 'D' to Downgrade - Enter 'C' to Cancel"

function ActionLogic {
  # action logic
  if ($Action -eq "U") {
    Upgrade
  }
  elseif ($Action -eq "D") {
    Downgrade
  }
  else {
    Cancel
  }
}

# initial run
ActionLogic

# initial confirmation prompt inside a conditional statement
if ($Action -eq "U") {
  $Confirmation = Read-Host "Are you sure with the changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
}
elseif ($Action -eq "D") {
  $Confirmation = Read-Host "Are you sure with the changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
}
else {
  Cancel
}

function ConfirmationLogic {
  if ($Confirmation -eq "Y") {
    if ($Action -eq "U") {
      ProceedUpgrade
    }
    elseif ($Action -eq "D") {
      ProceedDowngrade
    }
  }
  elseif ($Confirmation -eq "R") {
    $Col5.Value = $PlanDefaultValue
  }
  else {
    Cancel
  }
}

# run confirmation
ConfirmationLogic

while ($Confirmation -eq "R") {
  $Action = Read-Host "Repeating...`nEnter 'U' to Upgrade - Enter 'D' to Downgrade - Enter 'C' to Cancel"
  ActionLogic

  # logical cancelation
  if ($Action -eq "U") {
    $Confirmation = Read-Host "Are you reaaallyyy sure with the changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
    ConfirmationLogic
  }
  elseif ($Action -eq "D") {
    $Confirmation = Read-Host "Are you reaaallyyy sure with the changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
    ConfirmationLogic
  }
  else {
    $Confirmation = "C"
    Cancel
  }
}

# automatically exits the terminal session
function AutoExitTimer {
  Write-Host "This terminal will automatically close after 5 seconds . . . . ." -ForegroundColor DarkRed
  
  $Timer = [Diagnostics.Stopwatch]::StartNew()
  
  $Timer.Start()
  
  while ($Timer.Elapsed.Seconds -le 5) {
    # wait for 5 seconds
  }
  
  Write-Host "Farewell!!!" -ForegroundColor Blue
  
  while ($Timer.Elapsed.Seconds -le 7) {
    # wait for another 2 seconds
  }

  $Timer.Stop()
}

# run AutoExit
AutoExitTimer

# run taskkill.exe to kill all excel.exe processes for smooth execution of this command
TaskKill /IM Excel.exe /F

# garbage collection
[GC]::Collect()

# this automatically kills the current powershell session
[Environment]::Exit(0)
