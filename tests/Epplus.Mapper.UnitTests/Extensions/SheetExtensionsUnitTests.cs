using Epplus.Mapper.Annotations;
using Epplus.Mapper.Extensions;
using NUnit.Framework;
using OfficeOpenXml;
using System;
using System.IO;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class MyDto
    {
        [Cell("A1")]
        public int ColumnInt { get; set; }

        [Cell("B1")]
        public decimal ColumnDecimal { get; set; }

        [Cell("C1")]
        public DateTime ColumnDateTime { get; set; }

        [Cell("D1")]
        public string ColumnString { get; set; }

        [Cell("E1")]
        public bool ColumnBoolean { get; set; }

    }

    public class SheetExtensionsUnitTests
    {
        protected virtual ExcelWorksheet CreateTestingSheet()
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            return sheet;
        }

        [Test]

        public void Test_ExcelWorksheet_ReadValidCell()
        {
            var fileInfo = new FileInfo($"{TestContext.CurrentContext.TestDirectory}\\data.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets[1];
                var myDto = sheet.Read<MyDto>(2);

                Assert.AreEqual(1, myDto.ColumnInt);
                Assert.AreEqual(20.33, myDto.ColumnDecimal);
                Assert.AreEqual(new DateTime(2020, 2, 20), myDto.ColumnDateTime);
                Assert.AreEqual("Hello String", myDto.ColumnString);
                Assert.AreEqual(true, myDto.ColumnBoolean);
            }
        }

        [Test]
        public void Test_ExcelWorksheet_ReadInValidCell()
        {
            var fileInfo = new FileInfo($"{TestContext.CurrentContext.TestDirectory}\\data.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets[1];
                var myDto = sheet.Read<MyDto>(3);

                Assert.AreEqual(0, myDto.ColumnInt);
                Assert.AreEqual(0, myDto.ColumnDecimal);
                Assert.Less(myDto.ColumnDateTime, new DateTime(2000, 1, 1));
                Assert.AreEqual("83242", myDto.ColumnString);
                Assert.AreEqual(false, myDto.ColumnBoolean);
            }
        }
    }
}
