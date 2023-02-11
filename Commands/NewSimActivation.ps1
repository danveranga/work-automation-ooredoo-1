<# Using this command (custom cmdlet) will create a new instace of data to the excel database. This will also automatically add the value of the 1st column (Date Requested/Date Last Modified) to the date it was created, it is recommended to put the specific details of the new sim to be activated on the Remarks Column. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n" -ForegroundColor DarkRed

<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "$(Get-Location)\Commands\Database\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)
$LastUsedRow = $MainSheet.UsedRange.Rows.Count
$LastUnusedRow = $LastUsedRow + 1

# all columns
$Col1 = $MainSheet.Cells($LastUnusedRow, 1)   # Date Requested/Date Last Modified
$Col2 = $MainSheet.Cells($LastUnusedRow, 2)   # ICCID
$Col3 = $MainSheet.Cells($LastUnusedRow, 3)   # Request Type
$Col4 = $MainSheet.Cells($LastUnusedRow, 4)   # Mobile Number
$Col5 = $MainSheet.Cells($LastUnusedRow, 5)   # Current Plan
$Col6 = $MainSheet.Cells($LastUnusedRow, 6)   # Plan Rate
$Col7 = $MainSheet.Cells($LastUnusedRow, 7)   # Category/Location
$Col8 = $MainSheet.Cells($LastUnusedRow, 8)   # Employee No./Department
$Col9 = $MainSheet.Cells($LastUnusedRow, 9)   # Name
$Col10 = $MainSheet.Cells($LastUnusedRow, 10) # Designation
$Col11 = $MainSheet.Cells($LastUnusedRow, 11) # Staff Grade
$Col12 = $MainSheet.Cells($LastUnusedRow, 12) # Request Completion Date
$Col13 = $MainSheet.Cells($LastUnusedRow, 13) # Remarks

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

  $CurrentPlan = Read-Host "Ooredoo Plan to Apply (A-H only)"

  $OoredooPlans = @("A", "B", "C", "D", "E", "F", "G", "H")

  while ($CurrentPlan -notin $OoredooPlans) {
    $CurrentPlan = Read-Host "Your Input Plan is Invalid. Select Ooredoo Plan Again to Apply (A-H only)"
  }

  switch ($CurrentPlan) {
    "A" { $Col5.Value = "A"; $Col6.Value = "58.50" }
    "B" { $Col5.Value = "B"; $Col6.Value = "90" }
    "C" { $Col5.Value = "C"; $Col6.Value = "90" }
    "D" { $Col5.Value = "D"; $Col6.Value = "110.50" }
    "E" { $Col5.Value = "E"; $Col6.Value = "130" }
    "F" { $Col5.Value = "F"; $Col6.Value = "135" }
    "G" { $Col5.Value = "G"; $Col6.Value = "195" }
    "H" { $Col5.Value = "H"; $Col6.Value = "360" }

    Default { $Col5.Value = ""; $Col6.Value = "" }
  }

  $CategoryLocation = Read-Host "Enter the Category or Location"
  $Col7.Value = $CategoryLocation

  $EmpID_Department = Read-Host "Employee ID or Department of Ooredoo User"
  $Col8.Value = $EmpID_Department

  $EmpName = Read-Host "Name of Ooredoo User"
  $Col9.Value = $EmpName

  $Designation = Read-Host "Designation of Ooredoo User"
  $Col10.Value = $Designation

  $StaffGrade = Read-Host "Staff Grade of $($EmpID_Department) - $($EmpName)"
  $Col11.Value = $StaffGrade

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

  $Remarks = Read-Host "Other Remarks (optional)"
  $Col13.Value = "$($CurrentDateTime) - Ooredoo Sim Requested with Plan $($CurrentPlan); $($Remarks)"
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
$Confirmation = Read-Host "Are you sure you want to proceed with this information provided? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."

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
    $Excel = $null  # release the process
  }
}

# run ConfirmFunc
ConfirmFunc

# repeat loop until proceed or cancel have been selected
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
