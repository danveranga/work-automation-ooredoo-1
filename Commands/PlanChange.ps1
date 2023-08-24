<# This command (custom cmdlet) is used for upgrading or downgrading the plan of the specified mobile number by automatically changing Column 5 & 6 on the Ooredoo Master File based on the specified parameters. This will also automatically change the value of the Column 1 (Date Requested/Date Last Modified) to the date it was modified, while the Column 13 will automatically generate new details for the specified action (upgrade or downgrade). #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n
Don't forget to enter this command line [ TaskKill /IM Excel.exe /F ] after manually cancelling this command or go to task manager and manually kill the process of Excel application.`n
Ignoring this could create an error in re-running this command or running other commands in particular.`n" -ForegroundColor DarkRed

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

# all columns
$Col1 = $MainSheet.Cells($RowIndex, 1)   # ColA - Date Requested/Date Last Modified
# Column 2 / B is not used in this command.
# Column 3 / C is not used in this command.
$Col4 = $MainSheet.Cells($RowIndex, 4)   # ColD - Mobile Number
$Col5 = $MainSheet.Cells($RowIndex, 5)   # ColE - Plan Letter
$Col6 = $MainSheet.Cells($RowIndex, 6)   # ColF - Plan Rate
$Col7 = $MainSheet.Cells($RowIndex, 7)   # ColG - Plan Name
# Column 8 / H is not used in this command.
# Column 9 / I is not used in this command.
# Column 10 / J is not used in this command.
$Col11 = $MainSheet.Cells($RowIndex, 11) # ColK - Sim Holder
# Column 12 / L is not used in this command.
# Column 13 / M is not used in this command.
$Col14 = $MainSheet.Cells($RowIndex, 14) # ColN - Staff Grade
# Column 15 / O is not used in this command.
$Col16 = $MainSheet.Cells($RowIndex, 16) # ColP - Request Completion Date
$Col17 = $MainSheet.Cells($RowIndex, 17) # ColQ - Remarks (Activity Log)
# Column 18 / R is not used in this command.

# date and time details used for remarks
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# default value of Column 5: Ooredoo Plan Letter
$PlanDefaultValue = $Col5.Value2
$AamaliPlanDefaultValue = $Col7.Value2

# save the changes Upgrade Function has done
function ProceedUpgrade {

  # this will record the history of upgrade changes
  if ([string]::IsNullorEmpty($Col17.Value2)) {
    $Col17.Value = "$($CurrentDateTime) - From Plan $($PlanDefaultValue) - $($AamaliPlanDefaultValue) Upgraded to Plan $($Col5.Value2) - $($Col7.Value2)"
  }
  else {
    $Column17Value = $Col17.Value2
    $Col17.Value = "$($Column17Value)`n$($CurrentDateTime) - From Plan $($PlanDefaultValue) - $($AamaliPlanDefaultValue) Upgraded to Plan $($Col5.Value2) - $($Col7.Value2)"
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
  if ([string]::IsNullorEmpty($Col17.Value2)) {
    $Col17.Value = "$($CurrentDateTime) - From Plan $($PlanDefaultValue) - $($AamaliPlanDefaultValue) Downgraded to Plan $($Col5.Value2) - $($Col7.Value2)"
  }
  else {
    $Column17Value = $Col17.Value2
    $Col17.Value = "$($Column17Value)`n$($CurrentDateTime) - From Plan $($PlanDefaultValue) - $($AamaliPlanDefaultValue) Downgraded to Plan $($Col5.Value2) - $($Col7.Value2)"
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
    "B" { $Col5.Value = "B"; $Col6.Value = "72"; $Col7.Value = "Aamali 90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "104"; $Col7.Value = "Aamali 130" }
    "D" { $Col5.Value = "D"; $Col6.Value = "120"; $Col7.Value = "Aamali 150" }
    "E" { $Col5.Value = "E"; $Col6.Value = "175"; $Col7.Value = "Aamali 250" }
    "F" { $Col5.Value = "F"; $Col6.Value = "325"; $Col7.Value = "Aamali 500" }

    Default { $Col5.Value = $PlanDefaultValue; Write-Host "Invalid Plan Input; No Changes have been Made`nYou either repeat or cancel then start again." -ForegroundColor DarkBlue }  # nothing to do; remains the same
  }

  # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
  $Col16.Interior.ColorIndex = 6
  for ($i = 1; $i -lt $LastUsedRow; $i++) {
    $Col16Value = "R-$($i)"
    if ($Mainsheet.Range("P2:P$($LastUsedRow)").Value2 -notcontains $Col16Value) {
      $Col16.Value = $Col16Value
      break
    }
    else {
      $Col16.Value = ""
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
    "A" { $Col5.Value = "A"; $Col6.Value = "50.05"; $Col7.Value = "Aamali 65" }
    "B" { $Col5.Value = "B"; $Col6.Value = "72"; $Col7.Value = "Aamali 90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "104"; $Col7.Value = "Aamali 130" }
    "D" { $Col5.Value = "D"; $Col6.Value = "120"; $Col7.Value = "Aamali 150" }
    "E" { $Col5.Value = "E"; $Col6.Value = "175"; $Col7.Value = "Aamali 250" }

    Default { $Col5.Value = $PlanDefaultValue; Write-Host "Invalid Plan Input; No Changes have been Made`nYou either repeat or cancel then start again." -ForegroundColor DarkBlue }  # nothing to do; remains the same
  }

  # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
  $Col16.Interior.ColorIndex = 6
  for ($i = 1; $i -lt $LastUsedRow; $i++) {
    $Col16Value = "R-$($i)"
    if ($Mainsheet.Range("P2:P$($LastUsedRow)").Value2 -notcontains $Col16Value) {
      $Col16.Value = $Col16Value
      break
    }
    else {
      $Col16.Value = ""
    }
  }
}

# function that displays the necessary information of card holder
function InfoDisplay {
  function EligiblePlan {
    # plan eligibility values
    $NA = "Not Applicable as per the ACIFM Staff Grades and Benefits Section"
    $AB = "Plan A to Plan B only"
    $AD = "Plan A to Plan D only"
    $AE = "Plan A to Plan E only"
    $AF = "Plan A to Plan F"

    # switch statement conditions
    switch ($Col14.Value2) {
      "S1" { $NA }
      "S2" { $NA }
      "T1" { $NA }
      "T2" { $NA }
      "T3" { $NA }

      "S3" { $AB }

      "S4" { $AD }
      "T4A" { $AD }

      "T4B" { $AE }
      "T4C" { $AE }
      "M1A" { $AE }

      "M1B" { $AF }
      "M1C" { $AF }
      "M2A" { $AF }
      "M2B" { $AF }
      "M3" { $AF }
      "M4" { $AF }

      Default { "I don't know about that because the Grade Value is ' $($Col14.Value2) '. You may manually check the Ooredoo Master File.🤷🏼‍♂️" }
    }
    return
  }

  # display output format
  Write-Host "`n::::: OOREDOO MASTER FILE DETAILS :::::`n
Sim Holder:              $($Col11.Value2)
Mobile Number:           $($Col4.Value2)
Current Ooredoo Plan:    $PlanDefaultValue
Employee Grade:          $($Col14.Value2)
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

# confirmation prompt inside a conditional statement
if ($Action -eq "U") {
  $Confirmation = Read-Host "Are you sure with the upgrade changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
}
elseif ($Action -eq "D") {
  $Confirmation = Read-Host "Are you sure with the downgrade changes to mobile plan? Enter 'Y' to proceed, 'R' to repeat and 'C' to Cancel."
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
    $Col5.Value = $PlanDefaultValue  # needed to repeat the process with a clean slate
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
    $Confirmation = "C"  # needed to change the value of 'Confirmation' Variable to break from the while loop
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

# run taskkill.exe to kill all excel.exe processes for smooth execution of this command
TaskKill /IM Excel.exe /F

# run AutoExit
AutoExitTimer

# garbage collection
[GC]::Collect()

# this automatically kills the current powershell session
[Environment]::Exit(0)
