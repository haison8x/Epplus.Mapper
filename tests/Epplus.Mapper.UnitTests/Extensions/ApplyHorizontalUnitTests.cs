using Epplus.Mapper.Annotations;
using Epplus.Mapper.Extensions;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace Epplus.Mapper.UnitTests.Extensions
{
    public class ApplyHorizontalUnitTests : SheetExtensionsUnitTests
    {
        [Test]
        public void ApplyHorizontal_VerticalModelList_ModelShouldBeAppliedSuccesfully()
        {
            // Arrange
            var modelList = new List<HorizontalModel> { new HorizontalModel(), new HorizontalModel() };
            var sheet = CreateTestingSheet();
            var expectedDateTime = modelList[0].PropertyDateTime.ToString("{0:f}");

            // Act
            sheet.ApplyHorizontal(modelList);

            // Assert
            Assert.That(sheet.Cells[1, 1].Value, Is.EqualTo(modelList[0].PropertyString));
            Assert.That(((DateTime)sheet.Cells[2, 1].Value).ToString("{0:f}"), Is.EqualTo(expectedDateTime));
            Assert.That(sheet.Cells[1, 2].Value, Is.EqualTo(modelList[1].PropertyString));
            Assert.That(((DateTime)sheet.Cells[2, 2].Value).ToString("{0:f}"), Is.EqualTo(expectedDateTime));
        }

        public class HorizontalModel
        {
            public HorizontalModel()
            {
                PropertyString = "A1";
                PropertyDateTime = DateTime.Now;
            }

            [Cell("A1")]
            public string PropertyString { get; set; }

            [Cell("A2")]
            public DateTime PropertyDateTime { get; set; }
        }
    }
}
