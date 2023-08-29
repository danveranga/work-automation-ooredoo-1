<# Using this command (custom cmdlet) will mark the specified mobile number (ooredoo sim) row on the Ooredoo Master File as "Inactive", automatically change the value of the 1st column (Date Requested/Date Last Modified) to the date it was deactivated and log the changes to Remarks (Activity Log) Column. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n
Don't forget to enter this command line [ TaskKill /IM Excel.exe /F ] after manually cancelling this command or go to task manager and manually kill the process of Excel application.`n
Ignoring this could create an error in re-running this command or running other commands in particular.`n" -ForegroundColor DarkRed

Write-Host "Warning: Please enter the exact mobile number that needs to be modified. If you enter an invalid value, this command will keep running and prompting you for the correct mobile number until it matches a record in the database entries.`n" -ForegroundColor Cyan

# mobile number initialization
[string]$Number = Read-Host "Enter the Mobile Number to be Deactivated"

<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "$(Get-Location)\Commands\Database\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)
$LastRow = $MainSheet.UsedRange.Rows.Count

# query (main)
$QueryNumber = $MainSheet.Range("D2:D$($LastRow)").Find($Number)

# mobile number validation loop
if ($Number.Length -eq 8) {
  while ($QueryNumber.Value2 -ne $Number) {
    [string][ValidateLength(8, 8)][ValidateNotNullOrEmpty()]$Number = Read-Host "Enter the Mobile Number Again"
    # modified query (if validation runs)
    $QueryNumber = $MainSheet.Range("D2:D$LastRow").Find($Number)
  }
}
else {
  while ($QueryNumber.Value2 -ne $Number) {
    [string][ValidateLength(8, 8)][ValidateNotNullOrEmpty()]$Number = Read-Host "Enter the Mobile Number Again"
    # modified query (if validation runs)
    $QueryNumber = $MainSheet.Range("D2:D$LastRow").Find($Number)
  }
}

# row indentifier
$RowIndex = $QueryNumber.Row

# all columns
$Col1 = $MainSheet.Cells($RowIndex, 1)   # ColA - Date Requested/Date Last Modified
$Col2 = $MainSheet.Cells($RowIndex, 2)   # ColB - ICCID
$Col3 = $MainSheet.Cells($RowIndex, 3)   # ColC - Request Type
$Col4 = $MainSheet.Cells($RowIndex, 4)   # ColD - Mobile Number
$Col5 = $MainSheet.Cells($RowIndex, 5)   # ColE - Plan Letter
# Column 6 / F is not used in this command.
$Col7 = $MainSheet.Cells($RowIndex, 7)   # ColG - Plan Name
$Col8 = $MainSheet.Cells($RowIndex, 8)   # ColH - Employee No. (Person Responsible for Sim Usage)
# Column 9 / I is not used in this command.
# Column 10 / J is not used in this command.
$Col11 = $MainSheet.Cells($RowIndex, 11) # ColK - Sim Holder
$Col12 = $MainSheet.Cells($RowIndex, 12) # ColL - User Designation
$Col13 = $MainSheet.Cells($RowIndex, 13) # ColM - User Department
$Col14 = $MainSheet.Cells($RowIndex, 14) # ColN - Staff Grade
$Col15 = $MainSheet.Cells($RowIndex, 15) # ColO - Current Employment Status
$Col16 = $MainSheet.Cells($RowIndex, 16) # ColP - Request Completion Date
$Col17 = $MainSheet.Cells($RowIndex, 17) # ColQ - Remarks (Activity Log)
$Col18 = $MainSheet.Cells($RowIndex, 18) # ColR - Sim Card Status

# date and time definitions
$CurrentDate = Get-Date -Format "dd-MMM-yyyy"
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# special patch for repeated remarks
$RemarksOriginalValue = $Col17.Value2

# function that displays the necessary information of sim holder
function PreInfoDisplay {
  # display initial output format
  Write-Host "`n::::: CURRENT OOREDOO MASTER FILE DETAILS :::::`n
Sim Holder:                 $($Col11.Value2)
Employee ID:                $($Col8.Value2)
User Department:            $($Col13.Value2)
User Designation:           $($Col12.Value2)
User Staff Grade:           $($Col14.Value2)
Current Employment Status:  $($Col15.Value2)

ICCID:                      $($Col2.Value2)
Request Type:               $($Col3.Value2)
Mobile Number:              $($Col4.Value2)
Current Ooredoo Plan:       $($Col5.Value2) - $($Col7.Value2)
Current Sim Card Status:    $($Col18.Value2)

`nRemarks:
$($Col17.Value2)`n" -ForegroundColor Magenta
}

# main function
function SimDeactivation {
  # date requested - default automatic value
  $Col1.Value = $CurrentDate

  # request completed - default automatic value
  $Col16.Value = $CurrentDate

  # deactivates the specified sim card / mobile number
  $Col18.Value = "INACTIVE"

  # records the changes in the activity log column
  if ([string]::IsNullorEmpty($Col17.Value2)) {
    $Col17.Value = "$($CurrentDateTime) - Sim Card Mobile Number have been Deactivated"
  }
  else {
    $Col17.Value = "$($RemarksOriginalValue)`n$($CurrentDateTime) - Sim Card Mobile Number have been Deactivated"
  }
  # additional remarks feature
  $DefaultCol17Value = $Col17.Value2  # default value of Column 17
  $NewAdditionalRemarks = Read-Host "Other Remarks"
  $Col17.Value = "$($DefaultCol17Value); $($NewAdditionalRemarks)"
}

# timer function
function Timer {
  Write-Host "This terminal will automatically close after 5 seconds . . . . ." -ForegroundColor DarkRed

  # timer
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

# function to save changes
function SaveMe {
  Write-Host "Saving the Changes . . ." -ForegroundColor Blue
  $Excel.DisplayAlerts = $false
  $Workbook.Save()  # saves the file
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # completed process prompt message
  $Message = "Successfully Deactivated. Changes have been saved."
  Write-Host $Message -ForegroundColor Green

  # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
  TaskKill /IM Excel.exe /F

  # garbage collection
  [GC]::Collect()

  # run timer
  Timer

  # this automatically kills the current powershell session
  [Environment]::Exit(0)
}

# exit function
function AutoExit {  
  # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
  $Excel.DisplayAlerts = $false
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
  TaskKill /IM Excel.exe /F

  # garbage collection
  [GC]::Collect()

  # run timer
  Timer

  # this automatically kills the current powershell session
  [Environment]::Exit(0)
}

## main process flow

# conditional invocation of main function with validation that cancels the process if the Sim Card Status is already inactive.
if ($Col18.Value2 -eq "INACTIVE") {
  Write-Host "`nThe mobile number you entered is already inactive. This operation is being cancelled.`n" -ForegroundColor Red
  AutoExit
}
else {
  PreInfoDisplay  # diplay the information first
  SimDeactivation  # run main function
  SaveMe  # save the changes
}
