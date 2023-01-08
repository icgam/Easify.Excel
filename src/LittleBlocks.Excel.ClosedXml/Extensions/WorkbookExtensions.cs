using LittleBlocks.Excel;

namespace LittleBlocks.Excel.ClosedXml.Extensions
{
    public static class WorkbookExtensions
    {
        public static void Protect(this IWorkbook workbook, string password)
        {
            foreach (var worksheet in workbook.Worksheets) worksheet.Protect(password);
        }
    }    
}
