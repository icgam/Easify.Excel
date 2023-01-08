LittleBlocks.Excel 
============

![Release](https://github.com/LittleBlocks/LittleBlocks.Excel/workflows/Release%20build%20on%20master/main/badge.svg) ![CI](https://github.com/LittleBlocks/LittleBlocks.Excel/workflows/CI%20on%20Branches%20and%20PRs/badge.svg) ![](https://img.shields.io/nuget/v/LittleBlocks.Excel.svg?style=flat-square)

The project contains a series of components which gives the user the ability to load and save Excel files (2007 and newer) and access or change rows and cells. It also provides a fluent mapping interface similar to the concepts for Automapper or CsvHelper to map the Excel files to entities.

## Get Started

Description of how to start working with library

### How to use

Add the extension package to your project(s)

```
Install-Package LittleBlocks.Excel.Extensions
``` 

or 

```
dotnet add package LittleBlocks.Excel.Extensions
```

Then you can add the following code to your startup to enable excel support in **ServiceCollection**

```c#
services.AddExcel();
```

it adds the following services to your application

```c#
serviceCollection.AddTransient<IExcelMapperBuilder, ExcelMapperBuilder>(); // Used for excel schema mapping
serviceCollection.AddTransient<IWorkbookLoader, WorkbookLoader>(); // Used for direct excel manipulation. Needed by previous service

```

Now you cane work with IWorkbookLoader to load the excel file from file or stream

```c#

var workbookLoader = serviceProvider.GetRequiredService<IWorkbookLoader>();
var workbook = workbook.Load("filename");

// Manipulate the file

workbook.Save("Another file");

```

Or you can `ExcelMapperBuilder` to map and load the excel into an entity

```c#

public sealed class PersonModel
{
    public string Name { get; set; }
    public string PersonName { get; set; }
    public string PersonNameContainsSpecialCharacter { get; set; }
    public decimal? AmountInCurrency { get; set; }
    public string Currency { get; set; }
    public decimal? Salary { get; set; }
    public DateTime? BirthDate { get; set; }
    public string CountryName { get; set; }
    public bool? Certified { get; set; }
    public int? CertificationId { get; set; }
    public int? CustomRowNumber { get; set; }
}

// ...

var builder = serviceProvider.GetService<IExcelMapperBuilder>();
var workbook = workbook.Load("filename");
var mapper = builder.Build(workbook);

var personList = mapper.Map<PersonModel>("Sheet1", opt =>
                opt.ForMember(m => m.CustomRowNumber, Resolve.ByValue("Custom Identification No"))).ToList()
                
```

> **Important Notes**: We are using WINDOWS specific drawing library System.Drawing.Common which is a port for GDI+ on Windows, this solution may not be the best, possibly look for alternatives.

### How to Engage, Contribute, and Give Feedback

Description of the steps or process to be a contributor to the project.

Some of the best ways to contribute are to try things out, file issues, join in design conversations,
and make pull-requests.

* [Be an active contributor](./docs/CONTRIBUTING.md): Check out the contributing page to see the best places to log issues and start discussions.
* [Roadmap](./docs/ROADMAP.md): The schedule and milestone themes for project.

## Reporting bugs and useful features

Security issues and bugs should be reported by creating the relevant features and bugs in issues sections

## Related projects

- [LittleBlocks](https://github.com/LittleBlocks/LittleBlocks)
- [LittleBlocks.Ef](https://github.com/LittleBlocks/LittleBlocks.Ef)
- [LittleBlocks.Templates](https://github.com/LittleBlocks/LittleBlocks.Templates)
- [CloseXml](https://github.com/closedxml)



