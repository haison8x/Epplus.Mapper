using Epplus.Mapper.Annotations;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Epplus.Mapper.Extensions
{
    public static class ExcelWorksheetExtensions
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

        public static void ApplyModel(this ExcelWorksheet sheet, ExcelWorksheet templateSheet, object model, int row)
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

        public static void ApplyVertical<T>(this ExcelWorksheet sheet, IEnumerable<T> models, int startRowIndex = 0)
        {
            ApplyVertical(sheet, sheet, models, startRowIndex);
        }

        public static void ApplyVertical<T>(this ExcelWorksheet sheet, ExcelWorksheet templateSheet, IEnumerable<T> models, int startRowIndex = 0)
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
                var rowIndex = startRowIndex == 0 ? cellAddress.Row : startRowIndex;
                foreach (var model in models)
                {
                    var value = property.GetValue(model);
                    var destinationCell = sheet.Cells[rowIndex, cellAddress.Column];
                    templateSheet.Cells[address].Copy(destinationCell);
                    destinationCell.Value = value;
                    rowIndex++;
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

        public static void AutoMergeRowsHaveSameValue(this ExcelWorksheet sheet, string address)
        {
            var excelAddress = new ExcelAddress(address);

            AutoMergeRowsHaveSameValue(sheet, excelAddress);
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

        public static T Read<T>(this ExcelWorksheet sheet, int row = 1)
        {
            var methods = typeof(ExcelWorksheet).GetMethods(BindingFlags.Public | BindingFlags.Instance);
            var getValueMethod = methods.First(m => m.IsGenericMethod && m.Name == "GetValue");

            var type = typeof(T);
            var model = (T)Activator.CreateInstance(type);
            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                var address = GetExcelCellAddress(property);
                if (string.IsNullOrEmpty(address))
                {
                    continue;
                }

                var cellAddress = new ExcelCellAddress(address);

                if (property.PropertyType == typeof(bool))
                {
                    var cellValue = sheet.GetValue<string>(row, cellAddress.Column);

                    var booleanValue = !string.IsNullOrWhiteSpace(cellValue)
                        && cellValue != "0"
                        && !cellValue.Equals("FALSE", StringComparison.CurrentCultureIgnoreCase);

                    property.SetValue(model, booleanValue);
                    continue;
                }

                var getValueMethodGeneric = getValueMethod.MakeGenericMethod(property.PropertyType);
                var excelValue = getValueMethodGeneric.Invoke(sheet, new object[] { row, cellAddress.Column });

                property.SetValue(model, excelValue);
            }

            return model;
        }

        public static IEnumerable<T> ReadAll<T>(this ExcelWorksheet sheet, int row = 1)
        {
            for (var i = row; i <= sheet.Dimension.End.Row; i++)
            {
                if (IsEmptyRow(sheet, i))
                {
                    continue;
                }

                yield return Read<T>(sheet, i);
            }
        }

        private static bool IsEmptyRow(this ExcelWorksheet sheet, int row = 1)
        {
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                //check if the cell is empty or not
                if (sheet.Cells[row, col].Value != null)
                {
                    return false;
                }
            }

            return true;
        }
    }
}