using OfficeOpenXml;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class SheetExtensionsUnitTests
    {
        protected virtual ExcelWorksheet CreateTestingSheet()
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            return sheet;
        }
    }
}
