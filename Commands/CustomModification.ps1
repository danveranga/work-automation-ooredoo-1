<# Using this command (custom cmdlet) will modify the specified row on the Ooredoo Master File and automatically change the value of the 1st column (Date Requested/Date Last Modified) to the date it was modified, it is recommended to always update the remarks by specifying what type of changes/modification occured. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n
Don't forget to enter this command line [ TaskKill /IM Excel.exe /F ] after manually cancelling this command or go to task manager and manually kill the process of Excel application.`n
Ignoring this could create an error in re-running this command or running other commands in particular.`n" -ForegroundColor DarkRed

Write-Host "Warning: Please enter the exact mobile number that needs to be modified. If you enter an invalid value, this command will keep running and prompting you for the correct mobile number until it matches a record in the database entries.`n" -ForegroundColor Cyan

# mobile number initialization
[string]$Number = Read-Host "Enter the Mobile Number to be Modified"

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
$Col6 = $MainSheet.Cells($RowIndex, 6)   # ColF - Plan Rate
$Col7 = $MainSheet.Cells($RowIndex, 7)   # ColG - Plan Name
$Col8 = $MainSheet.Cells($RowIndex, 8)   # ColH - Employee No. (Person Responsible for Sim Usage)
# Column 9 / I is not used in this command.
$Col10 = $MainSheet.Cells($RowIndex, 10) # ColJ - Department/Location/Station Responsible for Sim Usage
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

# timer used before exiting the terminal session
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

# pre auto exit function
function PreAutoExit {  
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

# post auto exit function
function PostAutoExit {
    # run taskkill.exe to kill all excel.exe processes for smooth execution of this command
    TaskKill /IM Excel.exe /F
  
    # garbage collection
    [GC]::Collect()
  
    # run timer
    Timer
  
    # this automatically kills the current powershell session
    [Environment]::Exit(0)
}

# function that displays the necessary information of sim holder
function InfoDisplay {
    # display output format
    Write-Host "`n::::: OOREDOO MASTER FILE DETAILS :::::`n
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

# diplay the information first
InfoDisplay

if ($Col18.Value2 -eq "INACTIVE") {
    Write-Host "`nThe mobile number you entered is already inactive. This operation is being cancelled.`n" -ForegroundColor Red
    PreAutoExit  # automatically exits the application
}
else {
    # main function
    function CM {

        param (
            # ICCID parameter
            [Parameter(Mandatory = $true)]
            $ICCID,
            # request type: mobile (default value) or internet
            [Parameter(Mandatory = $true)]
            $RequestType,
            # mobile number
            [Parameter(Mandatory = $true)]
            $MobileNumber,
            # current ooredoo plan
            [Parameter(
                Mandatory = $true,
                HelpMessage = "A-F only")]
            $OoredooPlan,
            # sim holder - specific employee number
            [Parameter(Mandatory = $true)]
            $Option1_SimHolder_EmpNo,
            # sim holder - department/location/station
            [Parameter(Mandatory = $true)]
            $Option2_SimHolder_DeptLocationStation,
            # specific department of sim holder
            [Parameter(Mandatory = $true)]
            $Department,
            # remarks
            [Parameter(Mandatory = $true)]
            $Remarks
        )

        # date requested - default automatic value
        $Col1.Value = $CurrentDate

        # cm logic
        if ($ICCID -eq "") {
            # $Col2.Value = $Col2.Value
            # nothing happens, this an unecessary method of doing this workflow, might be changed in the future
        }
        else {
            $Col2.Value = $ICCID
        }

        if ($RequestType -eq "") {
            $Col3.Value = "Mobile"  # default value
        }
        else {
            $Col3.Value = $RequestType
        }

        if ($MobileNumber -eq "") {
            # $Col4.Value = $Col4.Value
        }
        else {
            $Col4.Value = $MobileNumber
        }

        if ($OoredooPlan -eq "") {
            # $Col5.Value = $Col5.Value
            # $Col6.Value = $Col6.Value
            # $Col7.Value = $Col7.Value
        }
        elseif ($OoredooPlan -eq "A") {
            $Col5.Value = "A"
            $Col6.Value = "50.05"
            $Col7.Value = "Aamali 65"
        }
        elseif ($OoredooPlan -eq "B") {
            $Col5.Value = "B"
            $Col6.Value = "72"
            $Col7.Value = "Aamali 90"
        }
        elseif ($OoredooPlan -eq "C") {
            $Col5.Value = "C"
            $Col6.Value = "104"
            $Col7.Value = "Aamali 130"
        }
        elseif ($OoredooPlan -eq "D") {
            $Col5.Value = "D"
            $Col6.Value = "120"
            $Col7.Value = "Aamali 150"
        }
        elseif ($OoredooPlan -eq "E") {
            $Col5.Value = "E"
            $Col6.Value = "175"
            $Col7.Value = "Aamali 250"
        }
        elseif ($OoredooPlan -eq "F") {
            $Col5.Value = "F"
            $Col6.Value = "325"
            $Col7.Value = "Aamali 500"
        }
        else {
            Write-Host "Error on Ooredoo Plan Input Value. Please Input A-F Only! Repeat the Process!!!" -ForegroundColor DarkRed
        }

        if ($Option1_SimHolder_EmpNo -eq "") {
            # $Col8.Value = $Col8.Value
        }
        else {
            $Col8.Value = $Option1_SimHolder_EmpNo
        }

        if ($Option2_SimHolder_DeptLocationStation -eq "") {
            # $Col10.Value = $Col10.Value
        }
        else {
            $Col10.Value = $Option2_SimHolder_DeptLocationStation
        }

        if ($Department -eq "") {
            # $Col13.Value = $Col13.Value
        }
        else {
            $Col13.Value = $Department
        }

        # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
        $Col16.Interior.ColorIndex = 6
        for ($i = 1; $i -lt $LastRow; $i++) {
            $Col16Value = "R-$($i)"
            if ($Mainsheet.Range("P2:P$($LastRow)").Value2 -notcontains $Col16Value) {
                $Col16.Value = $Col16Value
                break
            }
            else {
                $Col16.Value = ""
            }
        }

        # this logic records history of changes which creates activity log
        if ($Remarks -eq "") {
            # $Col17.Value = $Col17.Value2
            Write-Host "Do not proceed without remarks value, please select the repeat option and try again!`nMake sure to add value on remarks!" -ForegroundColor Red
        }
        else {
            if ([string]::IsNullorEmpty($Col17.Value2)) {
                $Col17.Value = "$CurrentDateTime - $Remarks"
            }
            else {
                $Column17Value = $Col17.Value2
                $Col17.Value = "$Column17Value`n$CurrentDateTime - $Remarks"
            }
        }
    }

    function Proceed {

        # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
        $Col16.Interior.ColorIndex = 6
    
        $Workbook.Save()  # saves the file
        $Excel.Quit()  # close excel
        $Excel = $null  # release the process

        # completed process prompt message
        $Message = "Successfully Modified."
        Write-Host $Message -ForegroundColor Green
    
    }

    # run main function
    CM

    # modified information validation
    $Confirmation = Read-Host "Are you sure you want to proceed with this information? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."

    function ConfirmFunc {
        if ($Confirmation -eq "R") {
            $Col17.Value = $RemarksOriginalValue
            CM
        }
        elseif ($Confirmation -eq "Y") {
            Proceed
        }
        else {
            # this will cancel the whole process of this command and to make sure Excel File is always closed but not saved though
            $Excel.DisplayAlerts = $false
            $Excel.Quit()  # close excel
            $Excel = $null  # release the process
        }
    }

    # run ConfirmFunc
    ConfirmFunc

    # looping through 'ConfirmFunc' Function until 'proceed' or 'cancel' option have been selected
    while ($Confirmation -eq "R") {
        $Confirmation = Read-Host "Are you reaalllyyy sure you want to proceed with this information? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."
        # loop through this function
        ConfirmFunc
    }

    # exit the program
    PostAutoExit
}