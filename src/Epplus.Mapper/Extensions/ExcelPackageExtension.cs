using OfficeOpenXml;

namespace Epplus.Mapper.Extensions
{
    public static class ExcelPackageExtension
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static byte[] GetAsByteArraySafe(this ExcelPackage excelPackage, string password)
        {
            return string.IsNullOrEmpty(password)
                ? excelPackage.GetAsByteArray()
                : excelPackage.GetAsByteArray(password);
        }
    }
}
