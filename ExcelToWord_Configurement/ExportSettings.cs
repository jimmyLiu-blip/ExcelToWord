namespace ExcelToWord_Configurement
{
    /// <summary>
    /// 匯出設定類別
    /// 集中管理所有可配置的參數
    /// </summary>
    public class ExportSettings
    {
        /// <summary>Excel 來源檔案路徑</summary>
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        /// <summary>Word 輸出資料夾</summary>
        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        /// <summary>要匯出的命名範圍清單</summary>
        public string[] TargetNames { get; set; } = { "ACL_1", "ACLN_1" };

        /// <summary>從第幾張工作表開始處理</summary>
        public int StartSheetIndex { get; set; } = 7;

        /// <summary>圖片統一寬度（公分）</summary>
        public float ImageWidthCm { get; set; } = 15;

        /// <summary>每次操作後的延遲時間（毫秒）</summary>
        public int DelayMs { get; set; } = 100;
    }
}