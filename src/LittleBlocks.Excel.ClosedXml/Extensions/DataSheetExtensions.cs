using ClosedXML.Excel;
using LittleBlocks.Excel;
using LittleBlocks.Excel.ClosedXml;

namespace LittleBlocks.Excel.ClosedXml.Extensions
{
    public static class DataSheetExtensions
    {
        public static void Protect(this IDataSheet dataSheet, string password)
        {
            var worksheet = (dataSheet as DataSheet)?.Internal;
            worksheet?.Protect(password)
                .AllowElement(XLSheetProtectionElements.SelectEverything)
                .AllowElement(XLSheetProtectionElements.FormatEverything)
                .AllowElement(XLSheetProtectionElements.Sort)
                .AllowElement(XLSheetProtectionElements.AutoFilter);
        }
    }
}

