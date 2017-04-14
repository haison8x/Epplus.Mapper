using Epplus.Mapper.Extensions;
using NUnit.Framework;
using OfficeOpenXml;
using System;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class AutoMergeRowsHaveSameValueUnitTests : SheetExtensionsUnitTests
    {
        [Test]
        public void AutoMergeRowsHaveSameValue_ExcelAddressContainsTwoColumns_ShouldThrowArgumentOutOfRangeException()
        {
            // Arrange
            var excelAddress = new ExcelAddress(2, 1, 2, 2);
            var sheet = CreateTestingSheet();

            // Act
            var exception = Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AutoMergeRowsHaveSameValue(excelAddress));

            // Assert
            Assert.That(exception.Message, Is.EqualTo("excelAddress must contain 1 column only\r\nParameter name: excelAddress"));
        }

        [Test]
        public void AutoMergeRowsHaveSameValue_ExcelAddressIsOutOfRange_ShouldThrowArgumentOutOfRangeException()
        {
            // Arrange
            var excelAddress = new ExcelAddress(2, 1, 10, 1);
            var sheet = CreateTestingSheet();

            // Act
            var exception = Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AutoMergeRowsHaveSameValue(excelAddress));

            // Assert
            Assert.That(exception.Message, Is.EqualTo("excelAddress must not overlap sheet dimension\r\nParameter name: excelAddress"));
        }

        [Test]
        public void AutoMergeRowsHaveSameValue_ExcelAddressIsInRange_ShouldAutoMergeCorrectly()
        {
            // Arrange
            var excelAddress = new ExcelAddress(1, 1, 6, 1);
            var sheet = CreateTestingSheet();

            // Act
            sheet.AutoMergeRowsHaveSameValue(excelAddress);

            // Assert
            Assert.True(sheet.Cells["A1:A4"].Merge);
            Assert.True(sheet.Cells["A5:A6"].Merge);
            Assert.False(sheet.Cells["A1:A7"].Merge);
        }

        protected override ExcelWorksheet CreateTestingSheet()
        {
            var sheet = base.CreateTestingSheet();

            sheet.Cells[1, 1].Value = 1;
            sheet.Cells[2, 1].Value = 1;
            sheet.Cells[3, 1].Value = 1;
            sheet.Cells[4, 1].Value = 1;
            sheet.Cells[5, 1].Value = 2;
            sheet.Cells[6, 1].Value = 2;
            sheet.Cells[7, 1].Value = 5;

            sheet.Cells[1, 2].Value = 1;
            sheet.Cells[2, 2].Value = 1;
            sheet.Cells[3, 2].Value = 2;
            sheet.Cells[4, 2].Value = 2;
            sheet.Cells[5, 2].Value = 3;
            sheet.Cells[6, 2].Value = 3;
            sheet.Cells[7, 2].Value = 5;

            return sheet;
        }
    }
}
