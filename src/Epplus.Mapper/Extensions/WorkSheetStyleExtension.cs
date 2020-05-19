using OfficeOpenXml;
using System.Drawing;

namespace Epplus.Mapper.Extensions
{
    public static class WorkSheetStyleExtension
    {
        public static void HighlightBackground(this ExcelWorksheet worksheet, string address, string formula, Color color)
        {
            var excelAddress = new ExcelAddress(address);
            var formatting = worksheet.ConditionalFormatting.AddExpression(excelAddress);
            formatting.Style.Fill.BackgroundColor.Color = color;
            formatting.Formula = formula;
        }

        public static void HighlightFont(this ExcelWorksheet worksheet, string address, string formula, Color color)
        {
            var excelAddress = new ExcelAddress(address);
            var formatting = worksheet.ConditionalFormatting.AddExpression(excelAddress);
            formatting.Style.Font.Color.Color = color;
            formatting.Formula = formula;
        }
    }
}