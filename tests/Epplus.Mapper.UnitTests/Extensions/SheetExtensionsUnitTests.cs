using OfficeOpenXml;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class SheetExtensionsUnitTests
    {
        protected ExcelWorksheet CreateTestingSheet()
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            return sheet;
        }
    }
}
