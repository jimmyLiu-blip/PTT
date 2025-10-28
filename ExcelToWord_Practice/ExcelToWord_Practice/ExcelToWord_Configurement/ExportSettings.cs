using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord.Configuration
{
    public class ExportSettings
    {
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        public string[] TargetNames { get; set; } = { "ACL_1","ACLN_1" };

        public int StartSheetIndex { get; set; } = 7;

        public float WidthCm { get; set; } = 18;

        public int DelayMs { get; set; } = 200;
    }
}
