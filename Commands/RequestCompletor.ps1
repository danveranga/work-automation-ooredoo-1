<# Using this command (custom cmdlet) will update the value of column 12 - OoredooMasterFile.xlsx named 'Request Completion Date' - to the date it was executed. Column 13 or Remarks will also be updated with a new added value stating what changes had happen. #>

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
    if ($MainSheet.Cells($i, 12).Interior.ColorIndex -eq 6) {
      $Column12 = $MainSheet.Cells($i, 12).Value2
      $Column8 = $MainSheet.Cells($i, 8).Value2
      $Column9 = $MainSheet.Cells($i, 9).Value2
      $Column10 = $MainSheet.Cells($i, 10).Value2
      $Column5 = $MainSheet.Cells($i, 5).Value2
      Write-Host "# Request Number: $($Column12)" -ForegroundColor Cyan
      Write-Host "Details: $($Column8) - $($Column9) - $($Column10) (with Plan $($Column5))`n" -ForegroundColor DarkMagenta
    }
  }
}

# main function
function RequestCompletor {
  $RequestSelection = Read-Host "Select Request Number"  # this will contain the request number to be used for querying the specific row of the cell to modify

  if ([string]::IsNullorEmpty($RequestSelection)) {
    Write-Host "Error: The request number you specified is empty. Any changes will not be saved. Please run the command again." -ForegroundColor Red
    # automatically exits
    AutoExit
    # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
    TaskKill /IM Excel.exe /F
    # this automatically kills the current powershell session
    [Environment]::Exit(0)
  }
  elseif ($MainSheet.Range("L2:L$($LastUsedRow)").Value2 -notcontains $RequestSelection) {
    Write-Host "Error: The request number you specified is not existing. Any changes will not be saved. Please run the command again." -ForegroundColor Red
    # automatically exits
    AutoExit
    # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
    TaskKill /IM Excel.exe /F
    # this automatically kills the current powershell session
    [Environment]::Exit(0)
  }
  elseif ($MainSheet.Range("L2:L$($LastUsedRow)").Value2 -match $RequestSelection) {
    $QueryDetails = $MainSheet.Range("L2:L$($LastUsedRow)").Find($RequestSelection).Row  # contains the specific row of the cell/s to modify
    $CurrentPendingMobileNumberRow = $MainSheet.Cells($QueryDetails, 4)
    $CurrentPlanRow = $MainSheet.Cells($QueryDetails, 5)
    $CurrentDateCompletionRow = $MainSheet.Cells($QueryDetails, 12)
    $CurrentRemarksRow = $MainSheet.Cells($QueryDetails, 13)

    if ($CurrentPendingMobileNumberRow.Interior.ColorIndex -eq 6) {
      # used for new sim activation completion
      $DefineMobileNumber = Read-Host "Enter the Service Mobile Number Provided by Ooredoo"
      
      if ($DefineMobileNumber -match "[1-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]") {
        $CurrentPendingMobileNumberRow.Value = $DefineMobileNumber
        $CurrentDateCompletionRow.Value = $CurrentDate

        # records the changes in remakrs
        if ([string]::IsNullorEmpty($CurrentRemarksRow.Value2)) {
          $CurrentRemarksRow.Value = "$($CurrentDateTime) - Request was Completed; Service Number is $($DefineMobileNumber) with Plan $($CurrentPlanRow.Value2)"
        }
        else {
          $Column13Value = $CurrentRemarksRow.Value2
          $CurrentRemarksRow.Value = "$($Column13Value)`n$($CurrentDateTime) - Request was Completed; Service Number is $($DefineMobileNumber) with Plan $($CurrentPlanRow.Value2)"
        }

        # remove the color of the cell in column 4 and column 12
        $CurrentPendingMobileNumberRow.Interior.ColorIndex = 0
        $CurrentDateCompletionRow.Interior.ColorIndex = 0
      }
      else {
        Write-Host "The mobile number you entered is invalid!" -ForegroundColor Red
        # automatically exits
        AutoExit
        # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
        TaskKill /IM Excel.exe /F
        # this automatically kills the current powershell session
        [Environment]::Exit(0)
      }
    }
    else {
      # used for plan change and custom modification completion
      # leaves the mobile number as is but changes the value of column 12
      $CurrentDateCompletionRow.Value = $CurrentDate

      # records the changes in remakrs
      if ([string]::IsNullorEmpty($CurrentRemarksRow.Value2)) {
        $CurrentRemarksRow.Value = "$($CurrentDateTime) - Request was Completed with Plan $($CurrentPlanRow.Value2))"
      }
      else {
        $Column13Value = $CurrentRemarksRow.Value2
        $CurrentRemarksRow.Value = "$($Column13Value)`n$($CurrentDateTime) - Request was Completed with Plan $($CurrentPlanRow.Value2)"
      }

      # remove the color of the cell in column 12
      $CurrentDateCompletionRow.Interior.ColorIndex = 0
    }
  }
  # additional remarks feature
  $DefaultCol13Value = $MainSheet.Cells($($QueryDetails), 13).Value2  # defualt value of Column 13
  $NewAdditionalRemarks = Read-Host "Other Remarks"
  $MainSheet.Cells($($QueryDetails), 13).Value = "$($DefaultCol13Value); $($NewAdditionalRemarks)"
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

  # run timer
  Timer
}

# exit function
function AutoExit {  
  # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
  $Excel.DisplayAlerts = $false
  $Excel.Quit()  # close excel
  $Excel = $null  # release the process

  # run timer
  Timer
}

# check first if there are pending requests
$LRange = $MainSheet.Range("L2:L$($LastUsedRow)").Value2
if ($LRange -match "R-") {
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

# run taskkill.exe to kill all excel.exe processes for smooth execution of this command
TaskKill /IM Excel.exe /F

# garbage collection
[GC]::Collect()

# this automatically kills the current powershell session
[Environment]::Exit(0)
