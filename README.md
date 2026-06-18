# XlsxCommand

## Summary
PowerShell 7 Module with the following CmdLets: (1) CmdLet ```Export-WorksheetXlsx``` to create Excel XLSX format files, and (2) CmdLet ```Import-WorksheetXlsx``` to read from Excel XLSX format files.

## Installation
```
Install-Module -Name XlsxCommand
```

## Syntax of included CmdLets
### Export-WorksheetXlsx syntax
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

### Import-WorksheetXlsx syntax
```
Import-WorksheetXlsx [-ExcelXLSXFilePath] <string>
    -TabNames
    [<CommonParameters>]
    

Import-WorksheetXlsx [-ExcelXLSXFilePath] <string>
    [-TabName <string>]
    [-AsRawCellData]
    [-Header <string[]>]
    [-First <int>]
    [-Last <int>]
    [-Skip <int>]
    [-Schema <PSTypeName>]
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
This example writes the same values as the example #1 and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property relative position:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx -DataType Number,String,Number -Align Center,Left,Right
```

### 3. Create an Excel XLSX Worksheet with given cell formatting by property name.
This example writes the same values as the example #1 and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand
Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstThreeProcesses.xlsx -DataTypeMap @{Id='Number'; Name='String'; CPU='Number'} -AlignMap @{Id='Center'; Name='Left'; CPU='Right'}
```

### 4. Create an Excel XLSX Worksheet with multiple tabs by group.
This example writes the values of the DisplayName, ServiceType, and Status properties of the current services, grouped by Status, separated in different tabs by group, into a new Excel Worksheet XLSX file at the given path:
```
Import-Module XlsxCommand
$tabs = Get-Service | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } | group Status;
Export-WorksheetXlsx $home\Documents\ServicesByStatus.xlsx -Group $tabs
```

### 5. Create an Excel XLSX Worksheet with multiple tabs.
This example writes the same values and tabs as the example #3, the example #4, and also specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand
$serviceByStatus = Get-Service | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } | group Status;
$processes = Get-Process | Select-Object -First 3 | Select-Object Id,Name,CPU
$typemap = @{Id='Number'; Name='String'; CPU='Number'; Service='String'; Type='String'; Status='String'}
$alignmap = @{Id='Center'; Name='Left'; CPU='Right'; Service='Left'; Type='Center'; Status='Center'}
$processes | Export-WorksheetXlsx C:\config\Processes.xlsx -Group $serviceByStatus -DataTypeMap $typemap -AlignMap $alignmap
```

## Import-WorksheetXlsx usage examples
(*Pending...*)