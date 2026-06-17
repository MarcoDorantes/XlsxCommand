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
### 1. Create an Excel XLSX Worksheet with default cell formats.
This example writes the values of the ID, Name, and CPU properties of the first 3 current processes into a new Excel Worksheet XLSX file at the given path:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx
```

### 2. Create an Excel XLSX Worksheet with given cell formatting by property relative position.
This example writes the same values as the previous example and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property relative position:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx -DataType Number,String,Number -Align Center,Left,Right
```

### 3. Create an Excel XLSX Worksheet with given cell formatting by property name.
This example writes the same values as the previous example and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx -DataTypeMap @{Id='Number'; Name='String'; CPU='Number'} -AlignMap @{Id='Center'; Name='Left'; CPU='Right'}
```
