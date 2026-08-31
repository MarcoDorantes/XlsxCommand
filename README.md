# XlsxCommand

## Summary
`XlsxCommand` is a PowerShell 7 Module with the following cmdlets: (1) `Export-WorksheetXlsx` to create Excel Worksheet tabs and store them as files in Excel Workbook format (XLSX), and (2) `Import-WorksheetXlsx` to read Excel Worksheet tabs from files in Excel Workbook format (XLSX).

The cell data processing of both included cmdlets are based on a strict two-dimensional tabular data structure with rows and columns, i.e., a ‘*table*’. The first row is the header row. The next rows are data rows.

An Excel Worksheet tab is a two-dimensional **spread** array, not a strict tabular structure. Hence, the included cmdlets only work best if the cells of each Worksheet tab in an Excel Workbook are organized as a ‘*table*’.

For the case of `Import-WorksheetXlsx`, like the [PowerShell `Import-Csv` Cmdlet](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv#notes), the header row determines the number of columns and the column names. The column names are also the names of the properties of the output objects added to the PowerShell Pipeline. The header row is interpreted to be the column headers, unless you use the `Header` parameter to specify column headers. If any row has more values than the header row, the additional values are ignored. On the other hand, if the `Schema` parameter is used, then the names of the properties of the output objects added to the PowerShell Pipeline are determined by the provided type.

## Directed acyclic graph (DAG) of major external dependencies
By '*major*' we mean a dependency which provides most of the functionality for a feature.

By '*external*' we mean a dependency whose design evolution is self-governing and separated from `XlsxCommand`.
```
                                           XlsxCommand
                            -------------------------------------------
                             |                  |                    |
                             |                  |                    |
                             |                  V                    |
                             |     ExcelXLSXWorksheetValidator       |
                             |                  |                    |
                             |                  |                    |
                             V                  |                    V
                      ExcelXLSXWorksheetWriter  |  ExcelXLSXWorksheetReader
                                      |         |       |            |
                                      |         |       |            |
                                      V         V       V            V
                                    DocumentFormat.OpenXml       IllyumL2T.Core
```
`XlsxCommand` directly depends on [ExcelXLSXWorksheetValidator](https://www.nuget.org/packages/ExcelXLSXWorksheetValidator/) for the rule checking of the Open XML SDK Validator in [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/). Such component checks the rules related to Microsoft Office Open XML Spreadsheet document format.

### Export-WorksheetXlsx Cmdlet external dependencies
`Export-WorksheetXlsx` directly depends on [ExcelXLSXWorksheetWriter](https://www.nuget.org/packages/ExcelXLSXWorksheetWriter/) for both abstrations: individual two-dimensional tabular grid, and a set of two-dimensional tabular grids.

`Export-WorksheetXlsx` transitively depends on [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) for Microsoft Office Open XML Spreadsheet document format processing.

### Import-WorksheetXlsx Cmdlet external dependencies
`Import-WorksheetXlsx` directly depends on [ExcelXLSXWorksheetReader](https://www.nuget.org/packages/ExcelXLSXWorksheetReader/) for individual two-dimensional tabular grid, a set of two-dimensional tabular grids, and LINQ-to-XLSX abstrations.

`Import-WorksheetXlsx` transitively depends on [IllyumL2T.Core](https://www.nuget.org/packages/IllyumL2T.Core/) for cell value/field parsing in LINQ-to-XLSX processing.

`Import-WorksheetXlsx` transitively depends on [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/) for Microsoft Office Open XML Spreadsheet document format processing.

## Installation
### Installation prerequisites
The `XlsxCommand` module has been tested (installation and included cmdlets) on the following runtime environment:
1. [PowerShell 7.6.0](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.0)
2. [PowerShellGet 2.2.5](https://learn.microsoft.com/en-us/powershell/module/powershellget/?view=powershellget-2.x)

The following PowerShell variable and command display data to see if the minimum required versions are already present:
1. For PowerShell version: `$PSVersionTable.PSVersion`
2. For PowerShellGet version: `Get-Module -ListAvailable | ? Name -eq PowerShellGet | ft Name,Version`

### Installation location and scope
The installation process includes a location for the `XlsxCommand` module. Such location is determined by the `Scope` parameter of the `Install-Module` Cmdlet. The accessibility of the installed module is also determined by the value of that `Scope` parameter (accessible to all users of the computer or accessible only to the current user of the computer).
As the default value for the `Scope` parameter varies, checking  the related documentation is in order: [Install-Module -Scope parameter](https://learn.microsoft.com/en-us/powershell/module/powershellget/install-module#-scope)

### Installation process of the latest version for the first time
The following command does not require administrative permissions on the local machine. The use of its default parameter values may prompt for a confirmation as stated in the related documentation: [Install-Module description](https://learn.microsoft.com/en-us/powershell/module/powershellget/install-module?view=powershellget-2.x#description)
```
Install-Module -Name XlsxCommand
```

### Installation process to upgrade to the latest version
The following command does not require administrative permissions on the local machine:
```
Install-Module -Name XlsxCommand -Force
```

### Uninstall a previous version
The following command searches for a specific version of `XlsxCommand` module in the local machine. If found, then displays what would look like if you were to run the uninstallation of such specific version.
```
Get-InstalledModule XlsxCommand -RequiredVersion 4.4.0 | Uninstall-Module -WhatIf
```

The following command searches for a specific version of `XlsxCommand` module in the local machine. If found, then asks for confirmation to execute the uninstallation of such specific version.
```
Get-InstalledModule XlsxCommand -RequiredVersion 4.4.0 | Uninstall-Module -Confirm
```

## Syntax of included cmdlets
### Export-WorksheetXlsx syntax
```
Export-WorksheetXlsx [-DestinationExcelXLSXFilePath] <string>
    [-InputObject <psobject>]
    [-InputHashtable <OrderedHashtable>]
    [-Group <Object[]>]
    [-Tabular <Object[]>]
    [-DataType {String | Number | Date | Time | Boolean}]
    [-Align {Left | Center | Right}]
    [-DataTypeMap <OrderedHashtable>]
    [-AlignMap <OrderedHashtable>]
    [-SkipExcelValidator]
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
    [-SkipExcelValidator]
    [<CommonParameters>]
```

## Export-WorksheetXlsx usage examples
### 1. Create an Excel Workbook (XLSX) of one Worksheet tab with default cell formats.
This example writes the values of the ID, Name, and CPU properties of the first six current processes into a new Worksheet tab and stores them as a new Excel Workbook (XLSX) file at the given path:
```
Import-Module XlsxCommand

Get-Process | Select-Object -First 6 `
    | Select-Object Id,Name,CPU `
    | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx
```

### 2. Create an Excel Workbook (XLSX) of one Worksheet tab with given cell formatting by column relative position.
This example writes the same values as the [example #1](#1-create-an-excel-workbook-xlsx-of-one-worksheet-tab-with-default-cell-formats) and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property relative position:
```
Import-Module XlsxCommand

Get-Process | Select-Object -First 6 `
    | Select-Object Id,Name,CPU `
    | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx `
        -DataType Number,String,Number `
        -Align Center,Left,Right
```

### 3. Create an Excel Workbook (XLSX) of one Worksheet tab with given cell formatting by property name.
This example writes the same values as the [example #1](#1-create-an-excel-workbook-xlsx-of-one-worksheet-tab-with-default-cell-formats) and specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand

Get-Process | Select-Object -First 6 `
    | Select-Object Id,Name,CPU `
    | Export-WorksheetXlsx $home\Documents\FirstSixProcesses.xlsx `
        -DataTypeMap @{Id='Number'; Name='String'; CPU='Number'} `
        -AlignMap @{Id='Center'; Name='Left'; CPU='Right'}
```

### 4. Create an Excel Workbook (XLSX) with multiple Worksheet tabs by group.
This example writes the values of the DisplayName, ServiceType, and Status properties of the current services, grouped by Status, separated in different Worksheet tabs by group, into a new Excel Workbook (XLSX) file at the given path:
```
Import-Module XlsxCommand

$tabs = Get-Service `
    | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } `
    | group Status

Export-WorksheetXlsx $home\Documents\ServicesByStatus.xlsx -Group $tabs
```

### 5. Create an Excel XLSX Worksheet with multiple tabs.
This example writes the same values and tabs as the [example #3](#3-create-an-excel-workbook-xlsx-of-one-worksheet-tab-with-given-cell-formatting-by-property-name) plus the [example #4](#4-create-an-excel-workbook-xlsx-with-multiple-worksheet-tabs-by-group), and also specifies the cell data type and cell horizontal alignment for the corresponding data cells by property name:
```
Import-Module XlsxCommand

$serviceByStatus = Get-Service `
    | %{ [PSCustomObject]@{Service=$_.DisplayName; Type=$_.ServiceType; Status=$_.Status} } `
    | group Status

$processes = Get-Process | Select-Object -First 6 | Select-Object Id,Name,CPU

$typemap = @{Id='Number'; Name='String'; CPU='Number'; Service='String'; Type='String'; Status='String'}
$alignmap = @{Id='Center'; Name='Left'; CPU='Right'; Service='Left'; Type='Center'; Status='Center'}

$processes | Export-WorksheetXlsx C:\config\Processes.xlsx `
    -Group $serviceByStatus `
    -DataTypeMap $typemap `
    -AlignMap $alignmap
```

### 6. Create an Excel XLSX Worksheet with multiple tabs using PowerShell classes.
*Note*: See [example #8](#8-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes) for another, a bit simpler, way to achieve the same as this example (given the latest `XlsxCommand` version is installed and loaded).

This example writes the same values and tabs as the [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), and also specifies the cell data type and cell horizontal alignment for the data cells by .NET CLR attributes in the corresponding properties of PowerShell classes:
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

$processes | Export-WorksheetXlsx Processes.xlsx -Group $serviceByStatus
```

### 7. Create an Excel XLSX Worksheet with cells of Excel Date data type using a PowerShell class.
*Note*: See [example #9](#9-create-an-excel-xlsx-worksheet-with-cells-of-excel-date-data-type-using-a-powershell-class) for another, a bit simpler, way to achieve the same as this example (given the latest `XlsxCommand` version is installed and loaded).

This example writes the values of the `Name`, and `LastWriteTime` properties of found PDF files into a new Excel Workbook (XLSX) file at the given path. The cells in the column related to the `Modified` class property are written as values of Excel `Date` data type:

```
Import-Module XlsxCommand

class PDF
{
    [System.ComponentModel.DataAnnotations.DataType('String')]
    $FileName

    [System.ComponentModel.DataAnnotations.DataType('Date')]
    $Modified

    PDF($file)
    {
        $this.FileName = $file.Name
        $this.Modified = $file.LastWriteTime
    }
}

ls -Filter *.pdf `
    | %{ [PDF]::new($_) } `
    | Export-WorksheetXlsx pdffiles.xlsx
```

### 8. Create an Excel XLSX Worksheet with multiple tabs using PowerShell classes.
This example writes the same values and tabs as the [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), and also specifies the cell horizontal alignment for the data cells by .NET CLR attributes in the corresponding properties of PowerShell classes. In this example, the cell data type is derived from the declared data type of each property:

```
Import-Module XlsxCommand

class ProcessView
{
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    [int]$Id

    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Left')]
    [string]$Name

    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Right')]
    [System.Nullable[int]]$CPU

    ProcessView($process)
    {
        $this.Id = $process.Id
        $this.Name = $process.Name
        $this.CPU = $process.CPU
    }
}

class ServiceView
{
    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Left')]
    [string]$Service

    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    [string]$Type

    [System.ComponentModel.DataAnnotations.DisplayFormat(DataFormatString = 'Center')]
    [string]$Status

    ServiceView($service)
    {
        $this.Service = $service.DisplayName
        $this.Type = $service.ServiceType
        $this.Status = $service.Status
    }
}

$serviceByStatus = Get-Service | %{ [ServiceView]::new($_) } | group Status;
$processes = Get-Process | Select-Object -First 6 | %{ [ProcessView]::new($_) }

$processes | Export-WorksheetXlsx Processes.xlsx -Group $serviceByStatus
```

### 9. Create an Excel XLSX Worksheet with cells of Excel Date data type using a PowerShell class.
This example writes the values of the `Name`, and `LastWriteTime` properties of found PDF files into a new Excel Workbook (XLSX) file at the given path. The cells in the column related to the `Modified` class property are written as values of Excel `Date` data type:

```
Import-Module XlsxCommand

class PDF
{
    [string]$FileName

    [DateTime]$Modified

    PDF($file)
    {
        $this.FileName = $file.Name
        $this.Modified = $file.LastWriteTime
    }
}

ls -Filter *.pdf `
    | %{ [PDF]::new($_) } `
    | Export-WorksheetXlsx pdffiles.xlsx
```

## Import-WorksheetXlsx usage examples
### 10. Read all existing tab names from an Excel Workbook (XLSX), not just Worksheet tabs.
This example reads all Worksheet tab names found in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds an instance of `System.String` to the PowerShell Pipeline per tab name:
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -TabNames

Processes
Stopped
Running
```

### 11. Read first existing tab from an Excel XLSX Worksheet.
This example reads the header row and the first three data rows from the first Worksheet tab (by default if `TabName` parameter is not specified) in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per data row (the values of the properties are string representations of the cells in the row):
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -First 3

  Id Name     CPU
  -- ----     ---
2864 Process1 79.20
8208 Process2 48.56
6864 Process3 45.62
```

### 12. Read an existing tab by name from an Excel XLSX Worksheet.
This example reads the header row and the next three data rows from the Worksheet tab named '**Stopped**' in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per data row (the values of the properties are string representations of the cells in the row):
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -TabName 'Stopped' -First 3

Service                  Type              Status
-------                  ----              ------
Agent Activation Service Win32OwnProcess   Stopped
Witness Router Service   Win32ShareProcess Stopped
Tabular Writer Service   Win32OwnProcess   Stopped
```

### 13. Read data rows of cells with some null values from an Excel Worksheet tab.
This example reads the header row and the last three data rows from the first Worksheet tab (by default if `TabName` parameter is not specified) in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of `System.Management.Automation.PSCustomObject` to the PowerShell Pipeline per data row. The values of the properties are string representations of each cell value in the row; in this example, some of those string values are `null` (see values of the `CPU` property):
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\config\Processes.xlsx -Last 3

  Id Name     CPU
  -- ----     ---
4164 Process4 0.16
4436 Process5 
 748 Process6 
```

### 14. Read data rows of cells as instances of a PowerShell class from an Excel Worksheet tab.
This example ignores the header row and reads the first three data rows of the first Worksheet in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of a PowerShell class (which is a .NET reference type) to the PowerShell Pipeline per data row. The values of the properties are parsed from the corresponding values of each cell in the row by position into the corresponding property type. For the case of integer representations without separation characters for thousands, the values are parsed directly (see the `$ProcessID` property declaration):
```
Import-Module XlsxCommand

class ProcessView
{
    [int]$ProcessID
    [string]$ProcessName
}

Import-WorksheetXlsx C:\config\Processes.xlsx -First 3 -Schema [ProcessView]

ProcessID ProcessName
--------- -----------
     2864 Process1
     8208 Process2
     6864 Process3
```

### 15. Read cell values of a data row as numeric from an Excel Worksheet tab.
This example ignores the header row and reads the last three data rows of the first Worksheet in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of a PowerShell class (which is a .NET reference type) to the PowerShell Pipeline per data row. The values of the properties are parsed from the corresponding values of each cell in the row by position into the corresponding property type. For this case of floating-point numeric (`double`, *whose default value is zero*) representations, the parsing process (see [Illyum/l2t](https://github.com/Illyum/l2t)) of the property values requires a custom .NET CLR attribute as shown:
```
$ProcessViewSource = @'
public class ProcessView
{
  public int Id { get; set; }

  public string Name { get; set; }

  [IllyumL2T.Core.ParseBehavior(NumberStyle = System.Globalization.NumberStyles.Number)]
  public double CPU { get; set; }
}
'@
$module_path = Split-Path (Get-Module XlsxCommand | select -expand Path)
$assemblies = @(
    Join-Path $module_path 'IllyumL2T.Core.dll'
    Join-Path $module_path 'netstandard.dll'
)
Add-Type -TypeDefinition $ProcessViewSource -ReferencedAssemblies $assemblies

Import-WorksheetXlsx C:\config\Processes.xlsx -Last 3 -Schema [ProcessView]

   Id Name     CPU
   -- ----     ---
 4164 Process4 0.16
 4436 Process5 0.00
  748 Process6 0.00
```

### 16. Read cell values of a data row as nullable numeric from an Excel Worksheet tab.
This example ignores the header row and reads the last three data rows of the first Worksheet in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of a PowerShell class (which is a .NET reference type) to the PowerShell Pipeline per data row. The values of the properties are parsed from the corresponding values of each cell in the row by position into the corresponding property type. For this case of ***nullable*** floating-point numeric (`System.Nullable<double>`, *whose default value is null*) representations, the parsing process (see [Illyum/l2t](https://github.com/Illyum/l2t)) of the property values requires a custom .NET CLR attribute as shown:
```
$ProcessViewSource = @'
public class ProcessView
{
  public int Id { get; set; }

  public string Name { get; set; }

  [IllyumL2T.Core.ParseBehavior(NumberStyle = System.Globalization.NumberStyles.Number)]
  public double? CPU { get; set; }
}
'@
$module_path = Split-Path (Get-Module XlsxCommand | select -expand Path)
$assemblies = @(
    Join-Path $module_path 'IllyumL2T.Core.dll'
    Join-Path $module_path 'netstandard.dll'
)
Add-Type -TypeDefinition $ProcessViewSource -ReferencedAssemblies $assemblies

Import-WorksheetXlsx C:\config\Processes.xlsx -Last 3 -Schema [ProcessView]

   Id Name     CPU
   -- ----     ---
 4164 Process4 0.16
 4436 Process5 
  748 Process6 
```

### 17. Attempt to read cell data as the wrong property numeric type using a PowerShell class.
This example attempts to read a cell value (`45.62`) and to parse it as an integer for a class property (`$CPU`) and shows the warning message of the corresponding parse failure. For that, this example ignores the header row and reads the third data row of the first Worksheet in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of a PowerShell class (which is a .NET reference type) to the PowerShell Pipeline for such third data row.
```
Import-Module XlsxCommand

class ProcessView
{
    [int]$Id
    [string]$Name
    [int]$CPU
}

Import-WorksheetXlsx C:\config\Processes.xlsx -First 1 -Skip 2 -Schema [ProcessView]

WARNING: CPU: Unparsable System.Int32 >>> 45.62

  Id Name     CPU
  -- ----     ---
6864 Process3   0
```

### 18. Read Excel dates, represented as numeric values, in data rows as .NET DateTime property values from a Worksheet tab.
This example reads from the first Worksheet in an Excel Workbook (XLSX) file named `LogFile.xlsx`. Such file was created at Excel Online service with the *Blank workbook* template. The first Worksheet contains two columns with a header row. The second column was filled with simple date values typed as *\<month-number\>\<dash\>\<day-number\>*:

**Without explicit data type specification:**
```
Import-Module XlsxCommand

Import-WorksheetXlsx C:\logs\LogFile.xlsx

LogName   LogDate
-------   -------
NameLog1  46204
NameLog2  46209
NameLog3  46211
```

**With explicit data type specification:**
```
Import-Module XlsxCommand

class LogRecord
{
    [string]$LogName
    [DateTime]$LogDate
}

Import-WorksheetXlsx C:\logs\LogFile.xlsx -Schema [LogRecord]

LogName   LogDate
-------   -------
NameLog1  7/1/2026 12:00:00 AM
NameLog2  7/6/2026 12:00:00 AM
NameLog3  7/8/2026 12:00:00 AM
```

### 19. Skip parsing the marked columns and read data rows of cells as instances of a PowerShell class from an Excel Worksheet tab.
This example skips the columns corresponding to the properties marked with the .NET CLR attribute `System.ComponentModel.DataAnnotations.Schema.NotMapped`, ignores the header row, and reads the first three data rows of the first Worksheet in the Excel Workbook (XLSX) created in [example #5](#5-create-an-excel-xlsx-worksheet-with-multiple-tabs), or in [example #6](#6-create-an-excel-xlsx-worksheet-with-multiple-tabs-using-powershell-classes), and adds one instance of a PowerShell class (which is a .NET reference type) to the PowerShell Pipeline per data row:
```
Import-Module XlsxCommand

class ProcessView
{
    [System.ComponentModel.DataAnnotations.Schema.NotMapped()]$IdIgnored
    [string]$Name
}

Import-WorksheetXlsx C:\config\Processes.xlsx -First 3 -Schema [ProcessView]

IdIgnored Name
--------- ----
          Process1
          Process2
          Process3
```
