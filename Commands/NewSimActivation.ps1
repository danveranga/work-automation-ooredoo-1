<# Using this command (custom cmdlet) will create a new instace of data to the excel database. This will also automatically add the value of the 1st column (Date Requested/Date Last Modified) to the date it was created, it is recommended to put the specific details of the new sim to be activated on the Remarks Column. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n
Don't forget to enter this command line [ TaskKill /IM Excel.exe /F ] after manually cancelling this command or go to task manager and manually kill the process of Excel application.`n
Ignoring this could create an error in re-running this command or running other commands in particular.`n" -ForegroundColor DarkRed

<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "$(Get-Location)\Commands\Database\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)
$LastUsedRow = $MainSheet.UsedRange.Rows.Count
$LastUnusedRow = $LastUsedRow + 1

# all columns
$Col1 = $MainSheet.Cells($LastUnusedRow, 1)   # ColA - Date Requested/Date Last Modified
$Col2 = $MainSheet.Cells($LastUnusedRow, 2)   # ColB - ICCID
$Col3 = $MainSheet.Cells($LastUnusedRow, 3)   # ColC - Request Type
$Col4 = $MainSheet.Cells($LastUnusedRow, 4)   # ColD - Mobile Number
$Col5 = $MainSheet.Cells($LastUnusedRow, 5)   # ColE - Plan Letter
$Col6 = $MainSheet.Cells($LastUnusedRow, 6)   # ColF - Plan Rate
$Col7 = $MainSheet.Cells($LastUnusedRow, 7)   # ColG - Plan Name
$Col8 = $MainSheet.Cells($LastUnusedRow, 8)   # ColH - Employee No. (Person Responsible for Sim Usage)
# Column 9 / I is not used in this command.
$Col10 = $MainSheet.Cells($LastUnusedRow, 10) # ColJ - Department/Location/Station Responsible for Sim Usage
# Column 11 / K is not used in this command.
# Column 12 / L is not used in this command.
$Col13 = $MainSheet.Cells($LastUnusedRow, 13) # ColM - User Department
# Column 14 / N is not used in this command.
# Column 15 / O is not used in this command.
$Col16 = $MainSheet.Cells($LastUnusedRow, 16) # ColP - Request Completion Date
$Col17 = $MainSheet.Cells($LastUnusedRow, 17) # ColQ - Remarks (Activity Log)
$Col18 = $MainSheet.Cells($LastUnusedRow, 18) # ColR - Sim Card Status

# date and time definitions
$CurrentDate = Get-Date -Format "dd-MMM-yyyy"
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# main function
function NewOoredooAccount {
  # date requested - default automatic value for column 1
  $Col1.Value = $CurrentDate

  $ICCID = Read-Host "Enter the ICCID of the Ooredoo Sim"
  $Col2.Value = $ICCID

  $RequestType = Read-Host "Request Type (Mobile or Internet)"
  if ($RequestType -eq "Internet") {
    $Col3.Value = "Internet"
  }
  else {
    $Col3.Value = "Mobile"
  }

  $Col4.Interior.ColorIndex = 6

  $CurrentPlan = Read-Host "Ooredoo Plan to Apply (A-F only)"

  $OoredooPlans = @("A", "B", "C", "D", "E", "F")

  while ($CurrentPlan -notin $OoredooPlans) {
    $CurrentPlan = Read-Host "Your Input Plan is Invalid. Select a Valid Ooredoo Plan Again to Apply (A-F only)"
  }

  switch ($CurrentPlan) {
    "A" { $Col5.Value = "A"; $Col6.Value = "50.05"; $Col7.Value = "Aamali 65" }
    "B" { $Col5.Value = "B"; $Col6.Value = "72"; $Col7.Value = "Aamali 90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "104"; $Col7.Value = "Aamali 130" }
    "D" { $Col5.Value = "D"; $Col6.Value = "120"; $Col7.Value = "Aamali 150" }
    "E" { $Col5.Value = "E"; $Col6.Value = "175"; $Col7.Value = "Aamali 250" }
    "F" { $Col5.Value = "F"; $Col6.Value = "325"; $Col7.Value = "Aamali 500" }

    Default { $Col5.Value = ""; $Col6.Value = "" }
  }

  $EmpID = Read-Host "Enter the Employee No. of Person Responsible for Sim Usage"
  $Col8.Value = $EmpID

  Write-Host "Reminder: If Employee No. of Person Responsible for Sim Usage is not necessary, you should put their Department, Location or Station instead in the next field."
  $DeptLocStation = Read-Host "Enter the Department, Location or Station of the Ooredoo User/s"
  $Col10.Value = $DeptLocStation

  $UserDept = Read-Host "Enter the Department Name"
  $Col13.Value = $UserDept

  $Col16.Interior.ColorIndex = 6
  for ($i = 1; $i -lt $LastUsedRow; $i++) {
    # a loop used to create a unique request ID which will be used for 'Request Completor' command.
    $Col16Value = "R-$($i)"
    if ($Mainsheet.Range("P2:P$($LastUsedRow)").Value2 -notcontains $Col16Value) {
      $Col16.Value = $Col16Value
      break
    }
    else {
      $Col16.Value = ""
    }
  }

  $ActLog = Read-Host "Other Remarks to Add (optional)"
  $Col17.Value = "$($CurrentDateTime) - Ooredoo Sim Requested with Plan $($CurrentPlan) - $($Col7.Value2) with the rate of $($Col6.Value2) QAR; $($ActLog)"

  $Col18.Value = "ACTIVE"
}

# run main function
NewOoredooAccount

# save function
function Proceed {
  $Excel.DisplayAlerts = $false  # prevents excel application pop-ups
  $Workbook.Save()  # saves the file
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # completed process prompt message
  $Message = "Successfully Added."
  Write-Host $Message -ForegroundColor Green
}

# logical confirmation
$Confirmation = Read-Host "Are you sure you want to proceed with the information provided? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."

function ConfirmFunc {
  if ($Confirmation -eq "R") {
    NewOoredooAccount
  }
  elseif ($Confirmation -eq "Y") {
    Proceed
  }
  else {
    # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
    $Excel.DisplayAlerts = $false  # prevents excel application pop-ups
    $Excel.Quit()  # close excel
    $Excel = $null  # release the process in the memory
  }
}

# run ConfirmFunc
ConfirmFunc

# looping through 'ConfirmFunc' Function until 'proceed' or 'cancel' option have been selected
while ($Confirmation -eq "R") {
  $Confirmation = Read-Host "Are you reaalllyyy sure you want to proceed with this information provided? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."
  # loop through this function
  ConfirmFunc
}

# automatically exits the terminal session
function AutoExitTimer {
  Write-Host "This terminal will automatically exit after 5 seconds . . . . ." -ForegroundColor DarkRed
  
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
