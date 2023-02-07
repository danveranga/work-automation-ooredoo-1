<# same with NSA but modified to fast-forward the process of activating 45 sim cards #>

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
  $Col2.Value = "89974010022202014$($ICCID)"

  $Col3.Value = "Mobile"

  $Col4.Interior.ColorIndex = 6

  $Col5.Value = "C"; $Col6.Value = "90"

  $Col7.Value = "Staff"

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

  $Remarks = "45 MEP Staff Sim Request as per GM"
  $Col13.Value = "$($CurrentDateTime) - Ooredoo Sim Requested with Plan C; $($Remarks)"
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

Proceed

# run taskkill.exe to kill all excel.exe processes for smooth execution of this command
TaskKill /IM Excel.exe /F

# garbage collection
[GC]::Collect()