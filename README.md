Easify.Excel 
============

The project contains a series of components which gives the user the ability to load and save Excel files (2007 and newer) and access or change rows and cells. It also provides a fluent mapping interface similar to the concepts for Automapper or CsvHelper to map the Excel files to entities.

## Get Started

Description of how to start working with library

### How to use

Add the extension package to your project(s)

```
Install-Package Easify.Excel.Extensions
``` 

or 

```
dotnet add package Easify.Excel.Extensions
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

- [Easify](https://github.com/icgam/Easify)
- [Easify.Ef](https://github.com/icgam/Easify.Ef)
- [Easify.Templates](https://github.com/icgam/Easify.Templates)
- [CloseXml](https://github.com/closedxml)



