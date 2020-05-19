using Epplus.Mapper.Annotations;
using Epplus.Mapper.Extensions;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class ApplyVerticalUnitTests : SheetExtensionsUnitTests
    {
        [Test]
        public void ApplyVertical_VerticalModelList_ModelShouldBeAppliedSuccesfully()
        {
            // Arrange
            var modelList = new List<VerticalModel> { new VerticalModel(), new VerticalModel() };
            var sheet = CreateTestingSheet();
            var expectedDateTime = modelList[0].PropertyDateTime.ToString("{0:f}");

            // Act
            sheet.ApplyVertical(modelList);

            // Assert
            Assert.That(sheet.Cells[1, 1].Value, Is.EqualTo(modelList[0].PropertyString));
            Assert.That(((DateTime)sheet.Cells[1, 2].Value).ToString("{0:f}"), Is.EqualTo(expectedDateTime));
            Assert.That(sheet.Cells[2, 1].Value, Is.EqualTo(modelList[1].PropertyString));
            Assert.That(((DateTime)sheet.Cells[2, 2].Value).ToString("{0:f}"), Is.EqualTo(expectedDateTime));
        }

        public class VerticalModel
        {
            public VerticalModel()
            {
                PropertyString = "A1";
                PropertyDateTime = DateTime.Now;
            }

            [Cell("A1")]
            public string PropertyString { get; set; }

            [Cell("B1")]
            public DateTime PropertyDateTime { get; set; }
        }
    }
}
