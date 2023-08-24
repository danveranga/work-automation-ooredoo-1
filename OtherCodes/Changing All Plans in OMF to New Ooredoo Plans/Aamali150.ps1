<# EXCEL - VBA OBJECTS #>
# excel objects initiation and invocation
$Excel = New-Object -ComObject Excel.Application  # initiates connection
$ExcelFilePath = "C:\Users\jcsar\OneDrive - Aktor Como Intercity Facilities Management\ACIFM Workspace\Ooredoo\[DRAFT] New Ooredoo Master List\OoredooMasterFile.xlsx"  # relative file path
$Workbook = $Excel.Workbooks.Open($ExcelFilePath)
$MainSheet = $Workbook.Sheets(1)

# date and time details used for remarks
$CurrentDateTime = Get-Date -Format "dd-MMM-yyyy @HH:mm"

# all columns needed
$E = $MainSheet # Plan Letter 5
$F = $MainSheet # Plan Rate 6
$G = $MainSheet # Plan Name 7
$Q = $MainSheet # Remarks 17

for ($i = 2; $i -lt 373; $i++) {
  if ($MainSheet.Cells($i, 7).Value2 -eq "Aamali 150") {
    $EE = $E.Cells($i, 5)
    $FF = $F.Cells($i, 6)
    $GG = $G.Cells($i, 7)
    $QQ = $Q.Cells($i, 17)

    # changes for remarks
    if ([string]::IsNullorEmpty($QQ.value2)) {
      $QQ.Value = "$($CurrentDateTime) - Old Plan Reference: Plan $($EE.value2) - $($GG.value2) with the rate of $($FF.Value2) QAR"
      Write-Host "shikakakakakakakakakakakakaakkakkakaakkakakakakakakaka"
    }
    else {
      $Column13Value = $QQ.Value2
      $QQ.Value = "$($Column13Value)`n$($CurrentDateTime) - Old Plan Reference: Plan $($EE.value2) - $($GG.value2) with the rate of $($FF.Value2) QAR"
    }

    #changes for the plan details
    $FF.Value = "120"
    $EE.Value = "D"
  }
  else {
    Write-Host $i
  }
}

$Workbook.Save()  # saves the file
$Excel.Quit()  # close excel
$Excel = $null  # release the process

# run taskkill.exe to kill all excel.exe processes for smooth execution of this command
TaskKill /IM Excel.exe /F

# garbage collection
[GC]::Collect()
