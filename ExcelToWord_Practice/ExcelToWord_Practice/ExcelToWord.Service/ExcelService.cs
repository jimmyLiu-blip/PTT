using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ExcelToWord.Service
{
    public class ExcelService:IExcelService
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _workbook;

        public ExcelService(string excelPath)
        {
            _excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
            };
            _workbook = _excelApp.Workbooks.Open(excelPath);
        }

        public Excel.Workbook Workbook => _workbook;

        public Excel.Range GetNameRange(Excel.Worksheet ws, string rangeName)
        { 
            Excel.Range range = null;

            try
            {
                range = _workbook.Names.Item(rangeName).RefersToRange;
            }
            catch 
            {
                try
                {
                    range = ws.Names.Item(rangeName).RefersToRange;
                }
                catch{}
            }
            return range;
        }

        public void Close()
        {
            try
            {
                _workbook?.Close(false);
                _excelApp?.Quit();

                if (_workbook != null)
                {
                    Marshal.FinalReleaseComObject(_workbook);
                }
                if (_excelApp != null)
                {
                    Marshal.FinalReleaseComObject(_excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch(Exception ex)
            {
                Console.WriteLine($"關閉excel時發生問題，{ex.Message}");
            }
        }
    }
}
