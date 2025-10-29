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
                range = _workbook.Names.Item(rangeName).RefersToRange;
            }
            catch
            {
                try
                {
                    range = ws.Names.Item(rangeName).RefersToRange;
                }
                catch { }
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
