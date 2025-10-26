using ExcelToWord_Configurement;
using System;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// 匯出流程協調器
    /// 負責協調 Excel 和 Word 服務，執行完整的匯出流程
    /// </summary>
    public class ExportCoordinator
    {
        private readonly ExportSettings _settings;
        private readonly IExcelService _excelService;
        private readonly IWordService _wordService;

        public ExportCoordinator(
            ExportSettings settings,
            IExcelService excelService,
            IWordService wordService)
        {
            _settings = settings;
            _excelService = excelService;
            _wordService = wordService;
        }

        public void Run()
        {
            // 建立輸出資料夾
            Directory.CreateDirectory(_settings.OutputFolder);

            Excel.Workbook workbook = _excelService.Workbook;

            // 處理每個工作表
            for (int i = _settings.StartSheetIndex; i <= workbook.Sheets.Count; i++)
            {
                ProcessWorksheet((Excel.Worksheet)workbook.Sheets[i]);
            }

            Console.WriteLine("\n🎉 全部完成！");

            // 清理資源
            _excelService.Close();
            _wordService.Quit();
        }

        /// <summary>處理單一工作表</summary>
        private void ProcessWorksheet(Excel.Worksheet ws)
        {
            Console.WriteLine($"\n▶ 處理工作表：{ws.Name}");

            foreach (string rangeName in _settings.TargetNames)
            {
                ProcessNamedRange(ws, rangeName);
            }
        }

        /// <summary>處理單一命名範圍</summary>
        private void ProcessNamedRange(Excel.Worksheet ws, string rangeName)
        {
            // 取得命名範圍
            Excel.Range range = _excelService.GetNamedRange(ws, rangeName);
            if (range == null)
            {
                Console.WriteLine($"⚠ 找不到命名範圍：{rangeName}（在 {ws.Name}）");
                return;
            }

            // 決定輸出檔案路徑
            string itemName = rangeName.Contains("_")
                ? rangeName.Split('_')[0]
                : rangeName;
            string wordPath = Path.Combine(_settings.OutputFolder, $"{itemName}.docx");

            // 開啟 Word 文件並插入圖片
            var doc = _wordService.OpenOrCreate(wordPath);
            _wordService.InsertRangePicture(doc, ws.Name, range, _settings.ImageWidthCm);
            _wordService.SaveAndClose(doc, wordPath);

            Console.WriteLine($"✅ 匯出 {rangeName} → {wordPath}");

            // 延遲確保 COM 操作完成
            Thread.Sleep(_settings.DelayMs);
        }
    }
}