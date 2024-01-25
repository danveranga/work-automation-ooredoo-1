<# same with RC but modified to fast-forward the process of completion of multiple sim card requests #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n" -ForegroundColor DarkRed

<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "$(Get-Location)\Commands\Database\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)
$LastUsedRow = $MainSheet.UsedRange.Rows.Count

# date and time definitions
$CurrentDate = Get-Date -Format "dd-MMM-yyyy"
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# function that shows the pending activities that requires completion; if nothing is pending, it will display on the screen that there are no pending activities/request to complete and thus will automatically close this command making it non-executable
function ShowPendingActivities {
  Write-Host "PENDING REQUESTS:`n" -ForegroundColor Blue
  for ($i = 1; $i -le $LastUsedRow; $i++) {
    if ($MainSheet.Cells($i, 16).Interior.ColorIndex -eq 6) {
      $Column16 = $MainSheet.Cells($i, 16).Value2  # request completion date (in this case the request number - if any)
      $Column8 = $MainSheet.Cells($i, 8).Value2  # emp id of person responsible
      $Column10 = $MainSheet.Cells($i, 10).Value2  # dept/loc/station of staff(s) responsible
      $Column11 = $MainSheet.Cells($i, 11).Value2  # sim holder
      $Column5 = $MainSheet.Cells($i, 5).Value2  # current plan letter
      $Column7 = $MainSheet.Cells($i, 7).Value2  # current plan name
      Write-Host "# Request Number: $($Column16)" -ForegroundColor Cyan
      Write-Host "Details: $($Column8) $($Column10) - $($Column11) (with Plan $($Column5) - $($Column7))`n" -ForegroundColor DarkMagenta
    }
  }
}

#############################################################################################################################################
# values to modify depending on the number of request to complete in one run
$MultipleRequesttoComplete = 1
$StartingRequestNumber = 0  # always decrement by one!!
#############################################################################################################################################

# main function
function RequestCompletor {
  for ($i = $StartingRequestNumber; $i -le $($MultipleRequesttoComplete + $StartingRequestNumber); $i++) {
    $RequestSelection = "R-$i"

    if ($MainSheet.Range("P2:P$($LastUsedRow)").Value2 -match $RequestSelection) {
      $QueryDetails = $MainSheet.Range("P2:P$($LastUsedRow)").Find($RequestSelection).Row  # contains the specific row of the cell/s to modify
      $CurrentPendingMobileNumberRow = $MainSheet.Cells($QueryDetails, 4)  # this code assumes that mobile number is already in the database
      $CurrentPlanLetterRow = $MainSheet.Cells($QueryDetails, 5)
      $CurrentPlanNameRow = $MainSheet.Cells($QueryDetails, 7)
      $CurrentDateCompletionRow = $MainSheet.Cells($QueryDetails, 16)
      $CurrentRemarksRow = $MainSheet.Cells($QueryDetails, 17)

      # records the changes in remakrs column
      if ([string]::IsNullorEmpty($CurrentRemarksRow.Value2)) {
        $CurrentRemarksRow.Value = "$($CurrentDateTime) - Request was Completed; Service Number is $($CurrentPendingMobileNumberRow.Value2) with Plan $($CurrentPlanLetterRow.Value2) - $($CurrentPlanNameRow.Value2)"
      }
      else {
        $Column13Value = $CurrentRemarksRow.Value2
        $CurrentRemarksRow.Value = "$($Column13Value)`n$($CurrentDateTime) - Request was Completed; Service Number is $($CurrentPendingMobileNumberRow.Value2) with Plan $($CurrentPlanLetterRow.Value2) - $($CurrentPlanNameRow.Value2)"
      }

      # removes the color of the cell and adds the current date to column 12
      $CurrentDateCompletionRow.Interior.ColorIndex = 0
      $CurrentDateCompletionRow.Value = $CurrentDate
    }
  }
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
  $Message = "Successfully Completed the Specified Request."
  Write-Host $Message -ForegroundColor Green

  # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
  TaskKill /IM Excel.exe /F

  # run timer
  Timer
}

# exit function
function AutoExit {  
  # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
  $Excel.DisplayAlerts = $false
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
  TaskKill /IM Excel.exe /F

  # run timer
  Timer
}

# check first if there are pending requests
$PColumnRange = $MainSheet.Range("P2:P$($LastUsedRow)").Value2
if ($PColumnRange -match "R-") {
  # show first the pending activities
  ShowPendingActivities
  # run the completor to modify the necessary columns
  RequestCompletor

  $Confirmation = Read-Host "Do you want to continue with this action? (Y/N)"
  if ($Confirmation -eq "Y") {
    # saves the changes
    SaveMe
  }
  else {
    AutoExit  # no changes will be saved
  }
}
else {
  # must exit as this cmdlet has not purpose if there are no pending activities to be completed
  Write-Host "There are no currently pending requests!`n" -ForegroundColor Green
  AutoExit  # no changes will be saved
}

# garbage collection
[GC]::Collect()
