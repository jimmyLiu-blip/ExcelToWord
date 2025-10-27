using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// Excel 操作服務實作類別
    /// 負責開啟 Excel、讀取命名範圍、關閉資源
    public class ExcelService : IExcelService
    {
        private readonly Excel.Application _excelApp;
        private readonly Excel.Workbook _workbook;

        public ExcelService(string excelPath)
        {
            _excelApp = new Excel.Application
            {
                Visible = false, // 看不見 excel視窗
                DisplayAlerts = false // 不要跳出提示
            };
            _workbook = _excelApp.Workbooks.Open(excelPath);
        }

        // 屬性（Property）」的簡化 Lambda 寫法
        // public Excel.Workbook Workbook
        // {
        // get { return _workbook; }
        // }
        public Excel.Workbook Workbook => _workbook;

        public Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName)
        {
            // 「預設 range 為 null，若成功找到命名範圍再覆蓋它，否則保持 null。」
            Excel.Range range = null;

            // 先找全域命名範圍，不限定在特定工作表
            // .RefersToRange：取得這個命名範圍實際指向的 Excel 儲存格範圍物件（Range）。
            // 從整個活頁簿層級中找到名為 rangeName 的命名範圍，
            // 並取出它所指向的那一塊儲存格（例如 $A$1:$I$12），存進變數 range。
            // 
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

        public void CloseWorkbook()
        {
            try
            {
                // _workbook?.Close(false); = 
                // if (_workbook != null) 
                // { _workbook.Close(false); }
                // .Close(false)：關閉活頁簿，不儲存變更
                // .Close(true)：關閉前先儲存變更
                // .Close()：預設等同 .Close(false)
                _workbook?.Close(false);
                // .Quit()呼叫後會把整個 Excel 程式退出背景執行
                // 但要記得釋放 COM 物件，否則 Excel.exe 仍留在背景中
                _excelApp?.Quit();

                // 釋放順序很重要
                // Marshal.ReleaseComObject(range);
                // Marshal.ReleaseComObject(worksheet);
                // Marshal.ReleaseComObject(workbook);
                // Marshal.ReleaseComObject(app);
                if (_workbook != null)
                    Marshal.ReleaseComObject(_workbook);
                if (_excelApp != null)
                    Marshal.ReleaseComObject(_excelApp);

                // 手動強制觸發 垃圾回收
                // C# 的 GC 會自動在記憶體不足時清理未使用物件
                // 但在處理 COM 物件時，有時需要立即清理釋放的 .NET Wrapper，
                // 因此我們會在釋放 COM 物件後加上：GC.Collect();
                // 立刻清理所有已經沒有參考的 .NET 物件，並釋放它們的記憶體。
                // GC.WaitForPendingFinalizers();通常跟在 GC.Collect();後面
                // 等待所有「終結器（finalizers）」執行完畢。
                // 有些物件會在釋放時執行「finalizer（析構子）」來清理底層資源，
                // 這些動作會在背景執行緒進行。
                // WaitForPendingFinalizers() 會暫停目前執行緒，
                // 直到所有終結器完成為止，確保釋放完全結束。
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"關閉 Excel 時發生問題：{ex.Message}");
            }
        }
    }
}

// COM（Component Object Model）是 Windows 的「舊式元件通訊技術」，
// 它讓不同語言（例如 C++、C#、VB）都能操控同一個應用程式（像 Excel）。
// Excel Interop 為什麼屬於 COM ?
// 因為執行：new Excel.Application();時
// C# 其實是在跟「Excel.exe 的 COM 伺服器」通訊。
// Excel 本體是由微軟 Office 提供的 COM 元件，C# 透過 Interop 桥接層來控制它。
// 因為 COM 元件不是純 .NET 物件，.NET 的垃圾回收（GC）無法自動清掉它，
// 所以你要手動釋放，否則 Excel.exe 會留在背景（永遠不關）。