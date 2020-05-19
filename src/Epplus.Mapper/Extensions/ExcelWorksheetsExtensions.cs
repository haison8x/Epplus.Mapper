using OfficeOpenXml;

namespace Epplus.Mapper.Extensions
{
    public static class ExcelWorksheetsExtensions
    {
        public static ExcelWorksheet CopyWithoutContent(this ExcelWorksheets sheets, string name, string newName)
        {
            var sheet = sheets.Copy(name, newName);
            var dimension = sheet.Dimension;

            sheet.Cells[dimension.Start.Row, dimension.Start.Column, dimension.End.Row, dimension.End.Column].Value = null;

            return sheet;
        }
    }
}
