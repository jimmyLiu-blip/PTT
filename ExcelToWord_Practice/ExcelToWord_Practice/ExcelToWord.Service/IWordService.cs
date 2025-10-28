using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToWord.Service
{
    public interface IWordService
    {
        Word.Document OpenOrCreate(string path);

        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm);

        void SaveAndClose(Word.Document doc, string path);

        void Close();
    }
}
