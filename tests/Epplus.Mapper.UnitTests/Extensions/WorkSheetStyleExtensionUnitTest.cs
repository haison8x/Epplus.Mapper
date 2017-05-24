using Epplus.Mapper.Extensions;
using NUnit.Framework;
using OfficeOpenXml;
using System.Drawing;

namespace Epplus.Mapper.UnitTests.Extensions
{
    [TestFixture]
    public class WorkSheetStyleExtensionUnitTest : SheetExtensionsUnitTests
    {
        [Test]
        public void HighlightBackgroup_ApplyToExcelSheet_ReturnAffectedAddress()
        {
            // Arrange
            var address = "A3:B6";
            var formula = "MOD(ROW(),2)>0";
            var color = Color.DarkRed;

            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightBackground(address, formula, color);
            var affectedAddress = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Address.Address;

            // Assert
            Assert.AreEqual("A3:B6", affectedAddress);
        }

        [Test]
        public void HighlightBackgroup_ApplyToExcelSheet_ReturnFormulaRule()
        {
            // Arrange
            var address = "A1:B8";
            var formula = "MOD(ROW(),2)>0";
            var color = Color.DarkRed;
            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightBackground(address, formula, color);
            var actualFormula = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Formula;

            // Assert
            Assert.AreEqual("MOD(ROW(),2)>0", actualFormula);
        }

        [Test]
        public void HighlightBackgroup_ApplyToExcelSheet_ReturnBackgroupColor()
        {
            // Arrange
            var address = "A4:B8";
            var formula = "MOD(ROW(),2)>0";
            var color = Color.SkyBlue;
            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightBackground(address, formula, color);
            var actualColor = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Style.Fill.BackgroundColor.Color;

            // Assert
            Assert.AreEqual(color, actualColor);
        }

        [Test]
        public void HighlightFont_SetCellAddress_ReturnAffectedCellAddress()
        {
            // Arrange
            var address = "A3";
            var formula = "A3>0";
            var color = Color.Red;
            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightFont(address, formula, color);
            var affectedAddress = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Address.Address;

            // Assert
            Assert.AreEqual("A3", affectedAddress);
        }

        [Test]
        public void HighlightFont_ApplyToExcelSheet_ReturnFormulaRule()
        {
            // Arrange
            var address = "A1:B8";
            var formula = "LEFT(E8,1)=\"-\"";
            var color = Color.DarkRed;
            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightFont(address, formula, color);
            var formattingFormula = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Formula;

            // Assert
            Assert.AreEqual(formula, formattingFormula);
        }

        [Test]
        public void HighlightFont_ApplyToExcelSheet_ReturnFontColor()
        {
            // Arrange
            var address = "A1:A8";
            var formula = "A1>0";
            var color = Color.Red;
            var sheet = CreateTestingSheet();

            // Act
            sheet.HighlightFont(address, formula, color);
            var formattingColor = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)(sheet.ConditionalFormatting[0]))
                .Style.Font.Color.Color;

            // Assert
            Assert.AreEqual(color, formattingColor);
        }

        protected override ExcelWorksheet CreateTestingSheet()
        {
            var sheet = base.CreateTestingSheet();

            sheet.Cells[1, 1].Value = 1;
            sheet.Cells[2, 1].Value = -1;
            sheet.Cells[3, 1].Value = 2;
            sheet.Cells[4, 1].Value = -2;
            sheet.Cells[5, 1].Value = 3;
            sheet.Cells[6, 1].Value = -3;
            sheet.Cells[7, 1].Value = 4;
            sheet.Cells[8, 1].Value = -4;

            sheet.Cells[1, 2].Value = "1(TEST)";
            sheet.Cells[2, 2].Value = "-1(TEST)";
            sheet.Cells[3, 2].Value = "2(TEST)";
            sheet.Cells[4, 2].Value = "-2(TEST)";
            sheet.Cells[5, 2].Value = "3(TEST)";
            sheet.Cells[6, 2].Value = "-3(TEST)";
            sheet.Cells[7, 2].Value = "4(TEST)";
            sheet.Cells[8, 2].Value = "-4(TEST)";

            return sheet;
        }
    }
}