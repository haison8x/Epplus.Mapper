using Epplus.Mapper.Annotations;
using Epplus.Mapper.Extensions;
using NUnit.Framework;
using System;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class ApplyModelUnitTests : SheetExtensionsUnitTests
    {
        [Test]
        public void ApplyModel_SimpleModel_ModelShouldBeAppliedSuccesfully()
        {
            // Arrange
            var model = new SimpleModel();
            var sheet = CreateTestingSheet();

            // Act
            sheet.ApplyModel(model);

            // Assert
            Assert.That(sheet.Cells[1, 2].Value, Is.EqualTo(model.PropertyString));
            Assert.That(sheet.Cells[1, 3].Value, Is.EqualTo(model.PropertyBool));
            Assert.That(sheet.Cells[1, 4].Value, Is.EqualTo(model.PropertyInt));
            Assert.That(sheet.Cells[1, 5].Value, Is.EqualTo(model.PropertyLong));
            Assert.That(sheet.Cells[1, 6].Value, Is.EqualTo(model.PropertyDecimal));
            Assert.That(sheet.Cells[1, 7].Value, Is.EqualTo(model.PropertyDouble));
            Assert.That(sheet.Cells[1, 8].Value, Is.EqualTo(model.PropertyDateTime));
        }

        [Test]
        public void ApplyModel_SimpleModel_AtAnotherRow_ModelShouldBeAppliedSuccesfully()
        {
            // Arrange
            var model = new SimpleModel();
            var sheet = CreateTestingSheet();

            // Act
            sheet.ApplyModel(sheet, model, 10);

            // Assert
            Assert.That(sheet.Cells[10, 2].Value, Is.EqualTo(model.PropertyString));
            Assert.That(sheet.Cells[10, 3].Value, Is.EqualTo(model.PropertyBool));
            Assert.That(sheet.Cells[10, 4].Value, Is.EqualTo(model.PropertyInt));
            Assert.That(sheet.Cells[10, 5].Value, Is.EqualTo(model.PropertyLong));
            Assert.That(sheet.Cells[10, 6].Value, Is.EqualTo(model.PropertyDecimal));
            Assert.That(sheet.Cells[10, 7].Value, Is.EqualTo(model.PropertyDouble));
            Assert.That(sheet.Cells[10, 8].Value, Is.EqualTo(model.PropertyDateTime));
        }

        public class SimpleModel
        {
            public SimpleModel()
            {
                PropertyString = "B1";
                PropertyBool = true;
                PropertyInt = 3;
                PropertyLong = 4;
                PropertyDecimal = 5;
                PropertyDouble = 6;
                PropertyDateTime = DateTime.Now;
            }

            [Cell("B1")]
            public string PropertyString { get; set; }

            [Cell("C1")]
            public bool PropertyBool { get; set; }

            [Cell("D1")]
            public int PropertyInt { get; set; }

            [Cell("E1")]
            public long PropertyLong { get; set; }

            [Cell("F1")]
            public decimal PropertyDecimal { get; set; }

            [Cell("G1")]
            public double PropertyDouble { get; set; }

            [Cell("H1")]
            public DateTime PropertyDateTime { get; set; }
        }
    }
}
