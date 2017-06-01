using Epplus.Mapper.Extensions;
using NUnit.CompareNetObjects;
using NUnit.Framework;
using OfficeOpenXml;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class ExcelPackageExtensionUnitTests
    {
        [Test]
        [TestCase("")]
        [TestCase(null)]
        public void GetAsByteArraySafe_EmptyPassword_ContentIsNotEncrypted(string password)
        {
            // Arrange
            var expectedExcelPackage = CreateTestExcelPackage();
            var expected = expectedExcelPackage.GetAsByteArray();

            // Act
            var excelPackage = CreateTestExcelPackage();
            var bytes = excelPackage.GetAsByteArraySafe(password);

            // Assert
            Assert.AreEqual(bytes.Length, expected.Length);
        }

        [Test]
        [TestCase("Any Not Empty Password")]
        public void GetAsByteArraySafe_NonEmptyPassword_ContentIsEncrypted(string password)
        {
            // Arrange
            var expectedExcelPackage = CreateTestExcelPackage();
            var expected = expectedExcelPackage.GetAsByteArray(password);

            // Act
            var excelPackage = CreateTestExcelPackage();
            var bytes = excelPackage.GetAsByteArraySafe(password);

            // Assert
            Assert.AreEqual(bytes.Length, expected.Length);
        }

        private ExcelPackage CreateTestExcelPackage()
        {
            var excelPackage = new ExcelPackage();
            var sheet = excelPackage.Workbook.Worksheets.Add("Content");
            sheet.Cells["A1"].Value = 0;

            return excelPackage;
        }
    }
}
