using OfficeOpenXml;

namespace Epplus.Mapper.Extensions
{
    public static class ExcelAddressBaseExtension
    {
        public static bool Contains(this ExcelAddressBase excelAddress, ExcelAddressBase another)
        {
            return excelAddress.Start.Row <= another.Start.Row
                && excelAddress.Start.Column <= another.Start.Column
                && excelAddress.End.Row >= another.End.Row
                && excelAddress.End.Column >= another.End.Column;
        }
    }
}
