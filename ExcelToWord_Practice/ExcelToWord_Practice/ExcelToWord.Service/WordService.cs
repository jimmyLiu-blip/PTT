using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelToWord.Service
{
    public class WordService : IWordService
    {
        private readonly Word.Application _wordApp;

        public WordService()
        {
            _wordApp = new Word.Application
            {
                Visible = true,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };

        }

    }
}
