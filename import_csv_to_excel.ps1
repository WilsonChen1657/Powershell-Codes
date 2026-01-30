param (
    [string]$inputCSV,
    [string]$outputXLSX
)
function Get-FilePathDialog ([string]$filterType = "All") {
    switch ($filterType) {
        "All" { $filter = "All files (*.*)|*.*" }
        "Excel" { $filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls" }
        Default { $filter = "$filterType files (*." + $filterType.ToLower() + ")|*." + $filterType.ToLower() }
    }
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $PWD.ProviderPath
    $dialog.Filter = $filter
    $dialog.ShowHelp = $true
    [void]$dialog.ShowDialog()
    return $dialog.FileName # Returns the full path of the selected file
}

#$inputCSV = "C:\Users\tpiwiche\Desktop\csvtoxlsx\net_data.csv"
#$outputXLSX = "C:\Users\tpiwiche\Desktop\csvtoxlsx\Net_data_mapping.xlsm"
### Set input and output path
if ([string]::IsNullOrEmpty($inputCSV)) {
    $inputCSV = Get-FilePathDialog -filterType "CSV"
    if ([string]::IsNullOrEmpty($inputCSV)) {
        throw "No input CSV file selected!"
    }
}
if ([string]::IsNullOrEmpty($outputXLSX)) {
    $outputXLSX = Get-FilePathDialog -filterType "Excel"
    if ([string]::IsNullOrEmpty($outputXLSX)) {
        throw "No input Excel file selected!"
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
### Read excel
$workbook = $excel.Workbooks.Open($outputXLSX)
$activeWorksheetIndex = 0
foreach ($worksheet in $workbook.Worksheets) {
    if ($worksheet.Name -eq "Main") {
        $activeWorksheetIndex = $worksheet.Index
    }
}
if ($activeWorksheetIndex -eq 0) {
    $activeWorksheetIndex = 1
}
### Activate Main worksheet
$workbook.Worksheets.Item($activeWorksheetIndex).Activate()
### Add a new worksheet before the active one
$worksheet = $workbook.Worksheets.Add()

### Build the QueryTables.Add command
### QueryTables does the same as when clicking "Data » From Text" in Excel
$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $worksheet.QueryTables.add($TxtConnector, $worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)

### Set the delimiter (, or ;) according to your regional settings
$query.TextFileOtherDelimiter = $Excel.Application.International(5)

### Set the format to delimited and text for every column
### A trick to create an array of 2s is used with the preceding comma
$query.TextFileParseType = 1
$query.TextFileColumnDataTypes = , 2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

### Execute & delete the import query
$query.Refresh()
$query.Delete()

# Activate Main worksheet
$workbook.Worksheets.Item($activeWorksheetIndex).Activate()

### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
# xlOpenXMLWorkbook             51	.xlsx	The default Excel workbook format (no macros).
# xlOpenXMLWorkbookMacroEnabled	52	.xlsm	Excel workbook with macros enabled.
# xlExcel8	                    56	.xls	The older Excel 97-2003 binary format.
# xlCSV	                         6	.csv	Comma-separated values file.
# xlTextWindows	                20	.txt    Tab-delimited text file.
$item = Get-Item $outputXLSX
switch ($item.Extension) {
    ".xlsx" { $format = 51 }
    ".xlsm" { $format = 52 }
    ".xls" { $format = 56 }
    ".csv" { $format = 6 }
    ".txt" { $format = 20 }
    Default { $format = 0 }
}
$Workbook.SaveAs($outputXLSX, $format)
$excel.Quit()
Write-Host "CSV import complete! $outputXLSX" -ForegroundColor Green
Exit

<#
$dir = [System.Environment]::CurrentDirectory
UtilityProgram\Out-EncryptedFile "$dir\import_csv_to_excel.ps1"
#>