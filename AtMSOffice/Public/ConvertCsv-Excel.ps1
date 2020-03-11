<#

#>
function funcConvertCsv-ExcelWithDelimiter {
    Param (
        <#Define locations and delimiter#>
        [Parameter(Mandatory=$true, Position=0)]
        [string]$csv,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$xlsx,
        [string]$delimiter = ','
    )
    $objExcel = New-Object -ComObject excel.application
    $objExcel.visible = $false
    $workbook = $objExcel.Workbooks.add()
    $s1 = $workbook.sheets | Where-Object {$_.name -eq 'Sheet1'}
    $s1.name = 'PowerShell Report'
    $TxtConnector = ('TEXT;' + $csv)
    $Connector = $s1.QueryTables.add($TxtConnector,$s1.Range('A1'))
    $Connector.name = 'Table1'
    $query = $s1.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $delimiter
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,1 * $s1.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    $query.Refresh()
    $query.Delete()
    $tablestyle = 1
    $table1=$workbook.ActiveSheet.ListObjects.add( $tablestyle,$workbook.ActiveSheet.UsedRange,0,1)
    $objExcel.ActiveSheet.UsedRange.EntireColumn.AutoFit()
    $objExcel.DisplayAlerts=$false
    $Workbook.SaveAs($xlsx,51)
    $objExcel.Quit()
}
function funcConvertCsv-Excel {
    Param (
        <#Define locations, default comma delimiter assumed#>
        [Parameter(Mandatory=$true, Position=0)]
        [string]$csv,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$xlsx
    )
    $objExcel = New-Object -ComObject excel.application 
    $objExcel.visible = $false
    $objExcel.DisplayAlerts=$false
    $workbook = $objExcel.workbooks.Open($csv)
    $tablestyle = 1
    $table1=$workbook.ActiveSheet.ListObjects.add( $tablestyle,$workbook.ActiveSheet.UsedRange,0,1)
    $objExcel.ActiveSheet.UsedRange.EntireColumn.AutoFit()
    $workbook.SaveAs($xlsx,51)
    $objExcel.Quit()
}
function prototypeConvertCsv-Excel {
    Param (
        <#example to show how simple it is#>
        [Parameter(Mandatory=$true, Position=0)]
        [string]$csv,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$xlsx
    )
    $objExcel = New-Object -ComObject excel.application 
    $objExcel.visible = $true
    $objExcel.DisplayAlerts=$false
    $workbook = $objExcel.workbooks.Open($csv)
    $workbook.SaveAs($xlsx,51)
    $objExcel.Quit()
}