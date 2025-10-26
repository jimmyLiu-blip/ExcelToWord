using System;

namespace ExcelToWord_Configurement
{
    /// 匯出設定類別
    /// 集中管理所有可配置的參數
    public class ExportSettings
    {
        /// Excel 來源檔案路徑
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        /// Word 輸出資料夾
        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        /// 要匯出的命名範圍清單
        public string[] TargetNames { get; set; } = { "ACL_1", "ACLN_1" };

        /// 從第幾張工作表開始處理
        public int StartSheetIndex { get; set; } = 7;

        /// 圖片統一寬度（公分）
        public float ImageWidthCm { get; set; } = 15;

        /// 每次操作後的延遲時間（毫秒）
        public int DelayMs { get; set; } = 150;

    }
}