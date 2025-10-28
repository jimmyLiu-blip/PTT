using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord.Service
{
    public interface IExcelService
    {
        Excel.Workbook Workbook { get; }

        Excel.Range GetNameRange(Excel.Worksheet ws, string rangeName);

        void Close();
    }
}
