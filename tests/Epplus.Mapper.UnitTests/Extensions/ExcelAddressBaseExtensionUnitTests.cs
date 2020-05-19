using Epplus.Mapper.Extensions;
using NUnit.Framework;
using OfficeOpenXml;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class ExcelAddressBaseExtensionUnitTests
    {
        private readonly ExcelAddress excelAddress = new ExcelAddress(5, 4, 10, 15);

        [Test]
        [TestCase(5, 4, 10, 15, ExpectedResult = true)]
        [TestCase(5, 6, 10, 15, ExpectedResult = true)]
        [TestCase(6, 4, 10, 15, ExpectedResult = true)]
        [TestCase(6, 6, 10, 15, ExpectedResult = true)]
        [TestCase(5, 4, 10, 14, ExpectedResult = true)]
        [TestCase(5, 4, 9, 15, ExpectedResult = true)]
        [TestCase(5, 4, 9, 14, ExpectedResult = true)]
        [TestCase(7, 7, 9, 8, ExpectedResult = true)]
        public bool Contains_AddressInside_ReturnsTrue(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            // Arrange
            var another = new ExcelAddress(fromRow, fromColumn, toRow, toColumn);

            // Act & Assert
            return excelAddress.Contains(another);
        }

        [Test]
        [TestCase(4, 4, 10, 15, ExpectedResult = false)]
        [TestCase(5, 3, 10, 15, ExpectedResult = false)]
        [TestCase(3, 3, 10, 15, ExpectedResult = false)]
        [TestCase(5, 4, 10, 16, ExpectedResult = false)]
        [TestCase(5, 4, 11, 15, ExpectedResult = false)]
        [TestCase(5, 4, 91, 16, ExpectedResult = false)]

        [TestCase(2, 1, 16, 20, ExpectedResult = false)]
        public bool Contains_AddressOutside_ReturnsFalse(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            // Arrange
            var another = new ExcelAddress(fromRow, fromColumn, toRow, toColumn);

            // Act & Assert
            return excelAddress.Contains(another);
        }
    }
}
