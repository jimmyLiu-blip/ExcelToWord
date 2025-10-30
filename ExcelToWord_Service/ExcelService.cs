using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace ExcelToWord.Service
{
    public class ExcelService : IExcelService
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _workbook;

        public ExcelService(string path)
        {
            _excelApp = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
            };
            _workbook = _excelApp.Workbooks.Open(path);
        }

        public Excel.Workbook Workbook => _workbook;

        public Excel.Range GetRangeName(Excel.Worksheet ws, string rangeName)
        {
            Excel.Range range = null;

            try
            {
                // range = ws.Range[rangeName];                     直接在 ws 尋找 rangeName
                range = ws.Names.Item(rangeName).RefersToRange;  // 從 ws的點名簿去找 rangeName
            }
            catch
            {
                try
                {
                    var nameRange = _workbook.Names.Item(rangeName);
                    if (nameRange != null)
                        range = nameRange.RefersToRange;
                }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"找不到命名範圍：{ws.Name}!{rangeName}");
                    Console.ResetColor();
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
            catch (Exception ex)
            {
                Console.WriteLine($"Excel出現異常狀況無法關閉，{ex.Message}");
            }
        }
    }
}
