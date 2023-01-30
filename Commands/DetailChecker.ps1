<# This command (custom cmdlet) is used to check the details of the specified mobile number by automatically querying the data on the Ooredoo Master File. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n" -ForegroundColor DarkRed

Write-Host "Warning: Please enter the exact mobile number that you need to check or query. If you enter an invalid value, this command will keep running and prompt you for the correct mobile number until it matches a record in the database entries.`n" -ForegroundColor Cyan

function DetailChecker {

  # mobile number initialization
  [string]$Number = Read-Host "Enter the Mobile Number to Query"

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
  $Col2 = $MainSheet.Cells($RowIndex, 2)   # ICCID
  $Col3 = $MainSheet.Cells($RowIndex, 3)   # Request Type
  $Col4 = $MainSheet.Cells($RowIndex, 4)   # MSISDN (Mobile Number)
  $Col5 = $MainSheet.Cells($RowIndex, 5)   # Current Plan
  $Col7 = $MainSheet.Cells($RowIndex, 7)   # Category/Location
  $Col8 = $MainSheet.Cells($RowIndex, 8)   # Employee No./Department
  $Col9 = $MainSheet.Cells($RowIndex, 9)   # Name
  $Col10 = $MainSheet.Cells($RowIndex, 10) # Designation
  $Col11 = $MainSheet.Cells($RowIndex, 11) # Staff Grade
  $Col13 = $MainSheet.Cells($RowIndex, 13) # Remarks

  # function that displays the necessary information of card holder
  function InfoDisplay {
    # display output format
    Write-Host "`n::::: OOREDOO MASTER FILE DETAILS :::::`n
Sim Holder:              $($Col9.Value2)
Employee Details:        $($Col10.Value2) - $($Col8.Value2)
Employee Grade:          $($Col11.Value2)
Category/Location:       $($Col7.Value2)

ICCID:                   $($Col2.Value2)
Type:                    $($Col3.Value2)
Mobile Number:           $($Col4.Value2)
Current Ooredoo Plan:    $($Col5.Value2)

Remarks:
$($Col13.Value2)`n" -ForegroundColor Magenta
  }

  InfoDisplay

  # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
  $Excel.DisplayAlerts = $false
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process
}

function AutoExitTimer {
  Write-Host "This terminal will automatically close after 5 seconds . . . . .`n" -ForegroundColor DarkRed
  
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

function ExitApp {
  # fancy
  Write-Host "`nClosing . . . . . . " -ForegroundColor Red

  # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
  TaskKill /IM Excel.exe /F

  # run AutoExit
  AutoExitTimer

  # garbage collection
  [GC]::Collect()

  # this automatically kills the current powershell session
  [Environment]::Exit(0)
}

# main control flow
DetailChecker

$Confirmation = Read-Host "Need to query again? (Y/N)"

while ($Confirmation -eq "Y") {
  if ($Confirmation -eq "Y") {
    DetailChecker
  }
  else {
    ExitApp
  }
  $Confirmation = Read-Host "Need to query something again? (Y/N)"
}

ExitApp