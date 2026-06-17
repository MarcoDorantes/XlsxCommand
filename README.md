# XlsxCommand

## Summary
PowerShell 7 Module with the following CmdLets: (1) CmdLet ```Export-WorksheetXlsx``` to create Excel XLSX format files, and (2) CmdLet ```Import-WorksheetXlsx``` to read from Excel XLSX format files.

## Installation
```
Install-Module -Name XlsxCommand
```

## Export-WorksheetXlsx syntax
```
Export-WorksheetXlsx [-DestinationExcelXLSXFilePath] <string>
    [-InputObject <psobject>]
    [-InputHashtable <OrderedHashtable>]
    [-Group <Object[]>]
    [-DataType {String | Number | Date | Time}]
    [-Align {Left | Center | Right}]
    [-DataTypeMap <OrderedHashtable>]
    [-AlignMap <OrderedHashtable>]
    [<CommonParameters>]
```

## Export-WorksheetXlsx usage examples
This example writes the values of the ID, Name, CPU properties of the first 3 current processes into a new Excel Worksheet XLSX file at the given path:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx
```
