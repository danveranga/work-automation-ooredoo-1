<# Using this command (custom cmdlet) will modify the specified row on the Ooredoo Master File and automatically change the value of the 1st column (Date Requested/Date Last Modified) to the date it was modified, it is recommended to always update the remarks by specifying what type of changes/modification occured. #>

Write-Host "`nMANDATORY INSTRUCTION: MAKE SURE TO SAVE AND CLOSE ALL EXCEL FILES BEFORE PROCEEDING WITH THIS COMMAND!`n
To cancel this command, press CTRL + C and then exit the Terminal.`n" -ForegroundColor DarkRed

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
$Col1 = $MainSheet.Cells($RowIndex, 1)   # Date Requested/Date Last Modified
$Col2 = $MainSheet.Cells($RowIndex, 2)   # ICCID
$Col3 = $MainSheet.Cells($RowIndex, 3)   # Request Type
$Col4 = $MainSheet.Cells($RowIndex, 4)   # MSISDN (Mobile Number)
$Col5 = $MainSheet.Cells($RowIndex, 5)   # Current Plan
$Col6 = $MainSheet.Cells($RowIndex, 6)   # Plan Rate
$Col7 = $MainSheet.Cells($RowIndex, 7)   # Category/Location
$Col8 = $MainSheet.Cells($RowIndex, 8)   # Employee No./Department
$Col9 = $MainSheet.Cells($RowIndex, 9)   # Name
$Col10 = $MainSheet.Cells($RowIndex, 10) # Designation
$Col11 = $MainSheet.Cells($RowIndex, 11) # Staff Grade
$Col12 = $MainSheet.Cells($RowIndex, 12) # Request Completion Date
$Col13 = $MainSheet.Cells($RowIndex, 13) # Remarks

# date and time definitions
$CurrentDate = Get-Date -Format "dd-MMM-yyyy"
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# special patch for repeated remarks
$RemarksOriginalValue = $Col13.Value2

# function that displays the necessary information of sim holder
function InfoDisplay {
    # display output format
    Write-Host "`n::::: MASTER FILE DETAILS :::::`n
Sim Holder:               $($Col9.Value2)
Employee No/Department:   $($Col8.Value2)
Designation:              $($Col10.Value2)
ICCID:                    $($Col2.Value2)
Mobile Number:            $($Col4.Value2)
Type:                     $($Col3.Value2)
Current Ooredoo Plan:     $($Col5.Value2)
Category/Location:        $($Col7.Value2)
Employee Grade:           $($Col11.Value2)
`nRemarks:
$($Col13.Value2)`n" -ForegroundColor Magenta
}

# diplay the information first
InfoDisplay

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
        # ooredoo plan/current plan
        [Parameter(
            Mandatory = $true,
            HelpMessage = "A-H only")]
        $OoredooPlan,
        # category/location
        [Parameter(Mandatory = $true)]
        $Category_Location,
        # employee number or department
        [Parameter(Mandatory = $true)]
        $EmpNo_Department,
        # name
        [Parameter(Mandatory = $true)]
        $Name,
        # designation
        [Parameter(Mandatory = $true)]
        $Designation,
        # grade
        [Parameter(Mandatory = $true)]
        $EmployeeGrade,
        # remarks
        [Parameter(Mandatory = $true)]
        $Remarks
    )

    # date requested - default automatic value
    $Col1.Value = $CurrentDate

    # cm logic
    if ($ICCID -eq "") {
        # $Col2.Value = $Col2.Value
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
    }
    elseif ($OoredooPlan -eq "A") {
        $Col5.Value = "A"
        $Col6.Value = "58.50"
    }
    elseif ($OoredooPlan -eq "B") {
        $Col5.Value = "B"
        $Col6.Value = "90"
    }
    elseif ($OoredooPlan -eq "C") {
        $Col5.Value = "C"
        $Col6.Value = "90"
    }
    elseif ($OoredooPlan -eq "D") {
        $Col5.Value = "D"
        $Col6.Value = "110.50"
    }
    elseif ($OoredooPlan -eq "E") {
        $Col5.Value = "E"
        $Col6.Value = "130"
    }
    elseif ($OoredooPlan -eq "F") {
        $Col5.Value = "F"
        $Col6.Value = "135"
    }
    elseif ($OoredooPlan -eq "G") {
        $Col5.Value = "G"
        $Col6.Value = "195"
    }
    elseif ($OoredooPlan -eq "H") {
        $Col5.Value = "H"
        $Col6.Value = "360"
    }
    else {
        Write-Host "Error on Ooredoo Plan Input Value. Please Input A-H Only! Repeat the Process!!!" -ForegroundColor DarkRed
    }

    if ($Category_Location -eq "") {
        # $Col7.Value = $Col7.Value
    }
    else {
        $Col7.Value = $Category_Location
    }

    if ($EmpNo_Department -eq "") {
        # $Col8.Value = $Col8.Value
    }
    else {
        $Col8.Value = $EmpNo_Department
    }

    if ($Name -eq "") {
        # $Col8.Value = $Col8.Value
    }
    else {
        $Col9.Value = $Name
    }

    if ($Designation -eq "") {
        # $Col8.Value = $Col8.Value
    }
    else {
        $Col10.Value = $Designation
    }

    if ($Grade -eq "") {
        # $Col8.Value = $Col8.Value
    }
    else {
        $Col11.Value = $Grade
    }

    # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
    $Col12.Interior.ColorIndex = 6
    for ($i = 1; $i -lt $LastRow; $i++) {
        $Col12Value = "R-$($i)"
        if ($Mainsheet.Range("L2:L$($LastRow)").Value2 -notcontains $Col12Value) {
            $Col12.Value = $Col12Value
            break
        }
        else {
            $Col12.Value = ""
        }
    }

    # this logic records history of changes
    if ($Remarks -eq "") {
        # $Col13.Value = $Col13.Value2
    }
    else {
        if ([string]::IsNullorEmpty($Col13.Value2)) {
            $Col13.Value = "$CurrentDateTime - $Remarks"
        }
        else {
            $Column13Value = $Col13.Value2
            $Col13.Value = "$Column13Value`n$CurrentDateTime - $Remarks"
        }
    }
}

function Proceed {

    # highlights the request completion date indicating that it's currently pending for completion - this is for 'RequestCompletor Command' use case
    $Col12.Interior.ColorIndex = 6
    
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
        $Col13.Value = $RemarksOriginalValue
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

# repeat loop until proceed or cancel have been selected
while ($Confirmation -eq "R") {
    $Confirmation = Read-Host "Are you reaalllyyy sure you want to proceed with this information? Enter 'R' to repeat, 'Y' to proceed and 'C' to cancel."
    # loop through this function
    ConfirmFunc
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
