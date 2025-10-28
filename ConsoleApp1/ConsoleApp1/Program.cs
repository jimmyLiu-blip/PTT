using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToWordByGroup_FinalWorking
{
    class Program
    {
        static void Main()
        {
            string excelPath = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";
            string outputFolder = @"C:\Reports\WordOutputs_ByItem";
            string[] targetNames = { "ACL_1", "ACLN_1" };
            int startSheetIndex = 7;

            Directory.CreateDirectory(outputFolder);

            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            var workbook = excelApp.Workbooks.Open(excelPath);

            try
            {
                for (int i = startSheetIndex; i <= workbook.Sheets.Count; i++)
                {
                    var ws = (Excel.Worksheet)workbook.Sheets[i];
                    Console.WriteLine($"\n▶ 處理工作表：{ws.Name}");

                    foreach (var rangeName in targetNames)
                    {
                        Excel.Range range = null;
                        try
                        {
                            range = workbook.Names.Item(rangeName).RefersToRange;
                        }
                        catch
                        {
                            try { range = ws.Names.Item(rangeName).RefersToRange; } catch { }
                        }

                        if (range == null)
                        {
                            Console.WriteLine($"⚠ 找不到命名範圍：{rangeName}（在 {ws.Name}）");
                            continue;
                        }

                        // Word 檔名
                        string itemName = rangeName.Contains("_") ? rangeName.Split('_')[0] : rangeName;
                        string wordPath = Path.Combine(outputFolder, $"{itemName}.docx");

                        // 【關鍵 1】複製圖片
                        range.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                        Excel.XlCopyPictureFormat.xlPicture);

                        // 【關鍵 2】等待複製完成
                        System.Threading.Thread.Sleep(200);

                        // 【關鍵 3】開啟 Word 文件
                        Word.Document doc;
                        if (File.Exists(wordPath))
                        {
                            doc = wordApp.Documents.Open(wordPath);
                        }
                        else
                        {
                            doc = wordApp.Documents.Add();
                        }

                        // 【關鍵 4】啟用文件
                        doc.Activate();

                        // 【關鍵 5】使用 Selection 插入內容
                        wordApp.Selection.EndKey(Unit: Word.WdUnits.wdStory);
                        wordApp.Selection.TypeText($"【{ws.Name}】");
                        wordApp.Selection.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                        wordApp.Selection.TypeParagraph();

                        // 【關鍵 6】立即貼上
                        wordApp.Selection.Paste();
                        wordApp.Selection.TypeParagraph();

                        // 【關鍵 7】立即存檔
                        doc.SaveAs2(wordPath);
                        doc.Close(SaveChanges: false);

                        Console.WriteLine($"✅ 匯出 {rangeName} → {wordPath}");

                        // 釋放物件
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    }
                }

                Console.WriteLine("\n🎉 全部完成！");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"❌ 發生錯誤：{ex.Message}");
                Console.WriteLine($"詳細：{ex.StackTrace}");
                Console.ResetColor();
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                wordApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}