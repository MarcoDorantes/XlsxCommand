# XlsxCommand

## Summary
`XlsxCommand` is a PowerShell 7 Module with the following CmdLets: (1) `Export-WorksheetXlsx` to create files as Excel XLSX format, and (2) `Import-WorksheetXlsx` to read from files as Excel XLSX format.

The cell data processing of both included CmdLets are based on a strict two-dimensional tabular data structure with rows and columns, i.e., a ‘*table*’. The first row is the header row with the names for each column. The next rows are data rows.

An Excel Worksheet tab is a two-dimensional **spread** array, not a strict tabular structure. Hence, the included CmdLets only work best if the cells of each tab in the Excel Worksheet are organized as a ‘*table*’.

## Installation
### Installation prerequisites
The `XlsxCommand` module has been tested (installation and included CmdLets) on the following runtime environment:
1. [PowerShell 7.6.0](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.0)
2. [PowerShellGet 2.2.5](https://learn.microsoft.com/en-us/powershell/module/powershellget/?view=powershellget-2.x)

### Installation process
The installation process includes a location for the `XlsxCommand` module. Such location is determined by the `Scope` parameter of the `Install-Module` CmdLet. The accessibility of the installed module is determined by the value of that `Scope` parameter (accessible to all users of the computer or accessible only to the current user of the computer).
As the default value for the `Scope` parameter varies, checking  the related documentation is in order: [Install-Module -Scope parameter](https://learn.microsoft.com/en-us/powershell/module/powershellget/install-module#-scope)
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

Get-Process | Select-Object -First 6 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx
```

### 2. Create an Excel XLSX Worksheet with given cell formatting by property relative position.
This example writes the same values as the example #1 and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property relative position:
```
Import-Module XlsxCommand

Get-Process | Select-Object -First 6 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx -DataType Number,String,Number -Align Center,Left,Right
```

### 3. Create an Excel XLSX Worksheet with given cell formatting by property name.
This example writes the same values as the example #1 and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand

Get-Process | Select-Object -First 6 | Select-Object Id,Name,CPU | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx -DataTypeMap @{Id='Number'; Name='String'; CPU='Number'} -AlignMap @{Id='Center'; Name='Left'; CPU='Right'}
```

### 4. Create an Excel XLSX Worksheet with multiple tabs by group.
This example writes the values of the DisplayName, ServiceType, and Status properties of the current services, grouped by Status, separated in different tabs by group, into a new Excel Worksheet XLSX file at the given path:
```
Import-Module XlsxCommand

$tabs = Get-Service | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } | group Status;
Export-WorksheetXlsx $home\Documents\ServicesByStatus.xlsx -Group $tabs
```

### 5. Create an Excel XLSX Worksheet with multiple tabs.
This example writes the same values and tabs as the example #3, or the example #4, and also specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand

$serviceByStatus = Get-Service | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } | group Status;
$processes = Get-Process | Select-Object -First 6 | Select-Object Id,Name,CPU
$typemap = @{Id='Number'; Name='String'; CPU='Number'; Service='String'; Type='String'; Status='String'}
$alignmap = @{Id='Center'; Name='Left'; CPU='Right'; Service='Left'; Type='Center'; Status='Center'}
$processes | Export-WorksheetXlsx C:\config\Processes.xlsx -Group $serviceByStatus -DataTypeMap $typemap -AlignMap $alignmap
```

### 6. Create an Excel XLSX Worksheet with multiple tabs using PowerShell classes.
This example writes the same values and tabs as the example #5, and also specifies the cell data type and cell horizontal alignment for the data cells by CLR attributes in the corresponding properties of PowerShell classes:
```
Import-Module XlsxCommand

class ProcessView
{
    [System.ComponentModel.DataAnnotations.DataType('Number')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    $Id

    [System.ComponentModel.DataAnnotations.DataType('String')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Left')]
    $Name

    [System.ComponentModel.DataAnnotations.DataType('Number')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Right')]
    $CPU

    ProcessView($process)
    {
        $this.Id = $process.Id
        $this.Name = $process.Name
        $this.CPU = $process.CPU
    }
}

class ServiceView
{
    [System.ComponentModel.DataAnnotations.DataType('String')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Left')]
    $Service

    [System.ComponentModel.DataAnnotations.DataType('String')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    $Type

    [System.ComponentModel.DataAnnotations.DataType('String')]
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    $Status

    ServiceView($service)
    {
        $this.Service = $service.DisplayName
        $this.Type = $service.ServiceType
        $this.Status = $service.Status
    }
}

$serviceByStatus = Get-Service | %{ [ServiceView]::new($_) } | group Status;
$processes = Get-Process | Select-Object -First 6 | %{ [ProcessView]::new($_) }
$processes | Export-WorksheetXlsx C:\config\Processes.xlsx -Group $serviceByStatus
```

## Import-WorksheetXlsx usage examples
### 7. Read existing tab names from an Excel XLSX Worksheet.
This example reads the tab names found in the Excel XLSX Worksheet created in example #5, or in example #6, and adds an instance of `System.String` to the PowerShell Pipeline per tab name:
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -TabNames

Processes
Stopped
Running
```

### 8. Read first existing tab from an Excel XLSX Worksheet.
This example reads the header row and the first three data rows from the first tab (by default if `TabName` parameter is not specified) of the Excel XLSX Worksheet created in example #5, or in example #6, and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per row of cells (the values of the properties are string representations of the cells in the row):
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -First 3

  Id Name     CPU
  -- ----     ---
2864 Process1 79.20
8208 Process2 48.56
6864 Process3 45.62
```

### 9. Read an existing tab by name from an Excel XLSX Worksheet.
This example reads the header row and the next three data rows from the tab named '**Stopped**' of the Excel XLSX Worksheet created in example #5, or in example #6, and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per row of cells (the values of the properties are string representations of the cells in the row):
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -TabName 'Stopped' -First 3

Service                  Type              Status
-------                  ----              ------
Agent Activation Service Win32OwnProcess   Stopped
Witness Router Service   Win32ShareProcess Stopped
Tabular Writer Service   Win32OwnProcess   Stopped
```

### 10. Read data rows of cells with some null values from an Excel XLSX Worksheet.
This example reads the header row and the last three data rows from the first tab (by default if `TabName` parameter is not specified) of the Excel XLSX Worksheet created in example #5, or in example #6, and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per row of cells. The values of the properties are string representations of each cell value in the row; in this example, some of those string values are null:
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -Last 3

  Id Name     CPU
  -- ----     ---
4164 Process4 0.16
4436 Process5 
 748 Process6 
```
