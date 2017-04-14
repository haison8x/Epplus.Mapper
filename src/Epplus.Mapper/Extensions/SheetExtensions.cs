using Epplus.Mapper.Annotations;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Epplus.Mapper.Extensions
{
    public static class SheetExtensions
    {
        private static readonly Type CellAttributeType = typeof(CellAttribute);

        public static void ApplyModel(this ExcelWorksheet sheet, object model)
        {
            var properties = model.GetType().GetProperties();
            foreach (var property in properties)
            {
                var address = GetExcelCellAddress(property);
                if (!string.IsNullOrEmpty(address))
                {
                    var value = property.GetValue(model);
                    sheet.Cells[address].Value = value;
                }
            }
        }

        public static void ApplyModel(this ExcelWorksheet sheet, ExcelWorksheet templateSheet, int row, object model)
        {
            var properties = model.GetType().GetProperties();
            foreach (var property in properties)
            {
                var address = GetExcelCellAddress(property);
                if (!string.IsNullOrEmpty(address))
                {
                    var value = property.GetValue(model);
                    var column = new ExcelAddress(address).Start.Column;
                    var destinationCell = sheet.Cells[row, column];
                    templateSheet.Cells[address].Copy(destinationCell);
                    destinationCell.Value = value;
                }
            }
        }

        public static void ApplyVertical<T>(this ExcelWorksheet sheet, IEnumerable<T> models)
        {
            var type = typeof(T);
            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var address = GetExcelCellAddress(property);
                if (string.IsNullOrEmpty(address))
                {
                    continue;
                }

                var cellAddress = new ExcelCellAddress(address);
                var rowIndex = 0;
                foreach (var model in models)
                {
                    var value = property.GetValue(model);
                    var destinationCell = sheet.Cells[cellAddress.Row + rowIndex, cellAddress.Column];
                    sheet.Cells[address].Copy(destinationCell);
                    destinationCell.Value = value;
                    rowIndex++;
                }
            }
        }

        public static void ApplyHorizontal<T>(this ExcelWorksheet sheet, IEnumerable<T> models)
        {
            var type = typeof(T);
            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var address = GetExcelCellAddress(property);
                if (string.IsNullOrEmpty(address))
                {
                    continue;
                }

                var cellAddress = new ExcelCellAddress(address);
                var colIndex = 0;
                foreach (var model in models)
                {
                    var value = property.GetValue(model);
                    var destinationCell = sheet.Cells[cellAddress.Row, cellAddress.Column + colIndex];
                    sheet.Cells[address].Copy(destinationCell);
                    destinationCell.Value = value;
                    colIndex++;
                }
            }
        }

        public static string GetExcelCellAddress(PropertyInfo property)
        {
            var attribute = property.GetCustomAttributes(CellAttributeType, true)
                   .Cast<CellAttribute>()
                   .FirstOrDefault();
            return attribute == null
                ? string.Empty
                : attribute.Address;
        }

        public static void AutoMergeRowsHaveSameValue(this ExcelWorksheet sheet, ExcelAddress excelAddress)
        {
            ValidateExcelAddressWhenAutoMergeRow(sheet, excelAddress);

            var columnIndex = excelAddress.Start.Column;
            var currentIndex = excelAddress.Start.Row;
            while (currentIndex <= excelAddress.End.Row)
            {
                var nextIndex = currentIndex + 1;
                var currentValue = sheet.Cells[currentIndex, columnIndex].Value;
                while (nextIndex <= excelAddress.End.Row && sheet.Cells[nextIndex, columnIndex].Value.Equals(currentValue))
                {
                    nextIndex++;
                }

                if (nextIndex - 1 > currentIndex)
                {
                    sheet.Cells[currentIndex, columnIndex, nextIndex - 1, columnIndex].Merge = true;
                }

                currentIndex = nextIndex;
            }
        }

        private static void ValidateExcelAddressWhenAutoMergeRow(ExcelWorksheet sheet, ExcelAddress excelAddress)
        {
            if (excelAddress.Start.Column != excelAddress.End.Column)
            {
                var message = $"{nameof(excelAddress)} must contain 1 column only";
                throw new ArgumentOutOfRangeException(nameof(excelAddress), message);
            }

            if (!sheet.Dimension.Contains(excelAddress))
            {
                var message = $"{nameof(excelAddress)} must not overlap sheet dimension";
                throw new ArgumentOutOfRangeException(nameof(excelAddress), message);
            }
        }
    }
}