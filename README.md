# Epplus.Mapper

Convention-based mapper between strong typed object and Excel data via Epplus.  
This project comes up with a task of my work, I am using it a lot in my project. Feel free to file bugs or raise pull requests...

## Install from NuGet
In the Package Manager Console:

`PM> Install-Package Epplus.Mapper`

## Get strong-typed objects from Excel (XLS or XLSX)

```C#
var package = new ExcelPackage();
var sheet = package.Workbook.Worksheets.Add("Sheet1");
var modelList = new List<VerticalModel> { new VerticalModel(), new VerticalModel() };
sheet.ApplyVertical(modelList);
