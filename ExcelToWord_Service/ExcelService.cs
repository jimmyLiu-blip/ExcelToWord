using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// Excel 操作服務實作類別
    /// 負責開啟 Excel、讀取命名範圍、關閉資源
    /// </summary>
    public class ExcelService : IExcelService
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _workbook;

        public ExcelService(string excelPath)
        {
            _excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };
            _workbook = _excelApp.Workbooks.Open(excelPath);
        }

        public Excel.Workbook Workbook => _workbook;

        public Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName)
        {
            Excel.Range range = null;

            // 先找全域命名範圍
            try
            {
                range = _workbook.Names.Item(rangeName).RefersToRange;
            }
            catch
            {
                // 再找工作表層級的命名範圍
                try
                {
                    range = ws.Names.Item(rangeName).RefersToRange;
                }
                catch
                {
                    // 找不到則保持 null
                }
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
                if (_excelApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠ 關閉 Excel 時發生問題：{ex.Message}");
            }
        }
    }
}