# 🧭 ExcelToWord 專案架構筆記

---

## 📂 專案分層說明

| 專案名稱 | 分層角色 | 功能說明 |
|-----------|-----------|-----------|
| **ExcelToWord** | 🏁 主程式層 (Entry Layer) | 啟動應用程式、載入設定、建立服務並呼叫執行流程 |
| **ExcelToWord_Configurement** | ⚙️ 設定層 (Configuration Layer) | 儲存所有可調整的設定，如 Excel 路徑、命名範圍、延遲時間等 |
| **ExcelToWord_Service** | 🔧 服務層 (Service Layer) | 核心邏輯所在，負責操作 Excel 與 Word（含介面與實作） |
| **ExcelToWord_Models** | 📦 資料模型層 (Domain / Models Layer) | 定義資料結構，目前暫時未用，日後可擴充（如測試項、報表結果物件） |
| **ExcelToWord_Utilities** | 🧰 工具層 (Utilities Layer) | 放置通用工具（例如日誌、檔案路徑、字串格式化工具等），目前未使用但保留擴充彈性 |

---

# 🧩 ExportSettings.cs

> **Namespace：** `ExcelToWord_Configurement`  
> 匯出設定類別 — 集中管理所有可配置的參數

```csharp
using System;

namespace ExcelToWord_Configurement
{
    /// <summary>
    /// 匯出設定類別，集中管理所有可配置的參數
    /// </summary>
    public class ExportSettings
    {
        // 1️⃣ Excel 來源檔案路徑
        /// <summary>Excel 來源檔案路徑</summary>
        /// <remarks>
        /// 字串前面的 <c>@</c> 用來讓字串中的「反斜線」不需要再轉義。
        /// - 沒有 @ 時：你要打兩個反斜線 `\\` 才能表示一個真正的 `\`。
        /// - 有 @ 時：反斜線會被當成普通字元，更符合實際檔案路徑的寫法。
        /// </remarks>
        public string ExcelPath { get; set; } = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";

        // 2️⃣ Word 輸出資料夾
        /// <summary>Word 報表輸出的資料夾路徑</summary>
        public string OutputFolder { get; set; } = @"C:\Reports\WordOutputs_ByItem";

        // 3️⃣ 要匯出的命名範圍清單
        /// <summary>要匯出的命名範圍清單</summary>
        /// <remarks>
        /// <c>string[]</c> 代表「字串陣列」。
        /// 
        /// 以下兩種寫法等價：
        /// ```csharp
        /// string[] names = { "A", "B", "C" };
        /// string[] names = new string[] { "A", "B", "C" };
        /// ```
        /// 因為 C# 會自動補上型別宣告，這稱為「陣列初始化簡寫」。
        /// </remarks>
        public string[] TargetNames { get; set; } = { "ACL_1", "ACLN_1" };

        // 4️⃣ 開始處理的工作表索引
        /// <summary>從 Excel 活頁簿中第幾張工作表開始處理</summary>
        public int StartSheetIndex { get; set; } = 7;

        // 5️⃣ 貼入 Word 圖片的統一寬度
        /// <summary>貼到 Word 文件中圖片的統一寬度（單位：公分）</summary>
        public float ImageWidthCm { get; set; } = 15f; // 使用 15f 確保型別為 float

        // 6️⃣ 每次操作後的延遲時間（毫秒）
        /// <summary>每次操作 COM 物件後的延遲時間（毫秒）</summary>
        /// <remarks>
        /// 用於防止 COM 阻塞問題，特別是在頻繁操作 Office 應用程式時。
        /// </remarks>
        public int DelayMs { get; set; } = 150;

        /* 🧠 延伸範例：自訂 Getter / Setter（僅示範用途）
        private int _internalDelayMs = 200;

        public int DelayMs
        {
            get
            {
                // 每次取得都回傳固定值 150
                return 150;
            }
            set
            {
                // 當有人嘗試設定值時，執行自訂邏輯
                Console.WriteLine($"有人試圖設定 DelayMs = {value}，但我們忽略它。");
                _internalDelayMs = 200; // 不使用外部給的 value
            }
        }
        */
    }
}
```
---
# 🧩 IExcelService 介面筆記

---

## 📘 檔案結構與命名空間

```csharp
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// Excel 操作服務介面
    /// 定義所有 Excel 相關操作的契約
    /// </summary>
    public interface IExcelService
    {
        /// <summary>
        /// 取得工作簿（Excel 檔案）
        /// </summary>
        /// <remarks>
        /// - 這是一個唯讀屬性，提供目前開啟的 Excel 活頁簿給外部使用。  
        /// - Excel.Workbook：屬性的型別 → 代表一個 Excel 檔案（活頁簿）。  
        /// - Workbook：屬性的名稱。
        /// </remarks>
        Excel.Workbook Workbook { get; }

        /// <summary>
        /// 取得指定工作表的命名範圍
        /// </summary>
        /// <remarks>
        /// - Excel.Range：方法的回傳型別 → 表示會回傳一個範圍物件（Range）。  
        /// - Excel.Worksheet ws：Excel 裡的一張「工作表」，參數名稱 ws 是常見縮寫（worksheet）。  
        /// - string rangeName：名稱管理員中的「命名範圍」，例如「ACL_1」指向一塊儲存格區域。
        /// </remarks>
        Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName);

        /// <summary>
        /// 關閉 Excel 並釋放資源
        /// </summary>
        void Close();
    }
}

Excel.Application        →  整個 Excel 程式本體
└── Workbooks            →  活頁簿集合
    └── Workbook         →  單一 Excel 檔案
        └── Worksheets   →  工作表集合
            └── Worksheet →  單一工作表 (Sheet)
                └── Range →  儲存格或範圍
```
---
# 🧩 ExcelService — Excel 操作服務實作筆記
```csharp
本類別負責開啟 Excel、讀取命名範圍、並正確關閉與釋放資源。
使用 Microsoft.Office.Interop.Excel (COM) 來操作 Excel。

## 📦 Namespace 與 using 區塊

using System;

using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

using Excel = ...
為了避免 Word 也有相同命名空間而導致衝突，給 Excel 取別名。

using System.Runtime.InteropServices;
用來釋放 COM 物件 (Marshal.ReleaseComObject())

---
## 🧱 類別結構
namespace ExcelToWord_Service
{
    /// Excel 操作服務實作類別

    /// 負責開啟 Excel、讀取命名範圍、關閉資源

    public class ExcelService : IExcelService
}

## 🔧 欄位與建構子
private readonly Excel.Application _excelApp;
private readonly Excel.Workbook _workbook;

public ExcelService(string excelPath)
{
    _excelApp = new Excel.Application
    {
        Visible = false,      // 看不見 Excel 視窗（背景執行）
        DisplayAlerts = false // 不要跳出提示訊息
    };

    // 開啟指定的 Excel 檔案
    _workbook = _excelApp.Workbooks.Open(excelPath);
}

✏️ 說明

Visible = false → Excel 在背景執行，不會彈出視窗。

DisplayAlerts = false → 關閉「是否要儲存」等提示，避免程式被中斷。

_workbook：代表目前開啟的活頁簿。

## 🧩 Workbook 屬性

// Lambda 簡化屬性寫法

public Excel.Workbook Workbook => _workbook;

✏️ 說明

這是「唯讀屬性（Read-only Property）」的簡化形式。
等價於：

public Excel.Workbook Workbook
{
    get { return _workbook; }
}

讓外部可以讀取 _workbook，但不能修改。

## 🎯 取得命名範圍
public Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName)
{
    Excel.Range range = null; // 預設為 null，找不到時回傳 null

    try
    {
        // 先找全域命名範圍（整個活頁簿層級）
        range = _workbook.Names.Item(rangeName).RefersToRange;
    }
    catch
    {
        try
        {
            // 若全域找不到，再嘗試找工作表層級
            range = ws.Names.Item(rangeName).RefersToRange;
        }
        catch
        {
            // 找不到則保持 null
        }
    }

    return range;
}

✏️ 說明

Excel.Names：Excel 檔案的「名稱管理員 (Name Manager)」。

.Item(rangeName)：從名稱集合中取得指定名稱的命名範圍。

.RefersToRange：取得該名稱實際對應的儲存格區域（例如 $A$1:$I$12）。

若找不到，range 會保持 null，方便外部判斷。

## ❎ 關閉 Excel 並釋放資源

public void CloseWorkbook()
{

    try
    {

        // 關閉目前活頁簿

        _workbook?.Close(false);

        // 關閉 Excel 應用程式

        _excelApp?.Quit();

        // 釋放 COM 物件
        if (_workbook != null)
            Marshal.ReleaseComObject(_workbook);
        if (_excelApp != null)
            Marshal.ReleaseComObject(_excelApp);

        // 強制執行垃圾回收
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
    catch (Exception ex)

    {
        Console.WriteLine($"關閉 Excel 時發生問題：{ex.Message}");

    }
}

✏️ 詳細說明

_workbook?.Close(false)	關閉活頁簿但不儲存變更（false = 不儲存，true = 儲存）

_excelApp?.Quit()	關閉整個 Excel 應用程式（無參數，沒有 Quit(false)）

Marshal.ReleaseComObject()	手動釋放 COM 物件引用，避免 Excel.exe 殘留在背景

GC.Collect()	強制觸發垃圾回收，清除所有未使用物件

GC.WaitForPendingFinalizers()	等待所有終結器（finalizers）執行完畢，確保釋放完整

## 🧠 COM（Component Object Model）簡介

定義：Windows 的舊式元件通訊技術，允許不同語言（C++、C#、VB）共用相同應用程式（如 Excel）。

Excel Interop 為何屬於 COM？

var app = new Excel.Application();


這行其實是透過 COM 與「Excel.exe」溝通，建立一個背景 Excel 實例。

為什麼要手動釋放？

COM 物件不是純 .NET 物件，GC 無法自動回收它。

若不釋放，Excel.exe 會留在背景（永遠不關）。

## 🧭 整體流程圖（概念順序）
ExcelService 建構子

    ↓
new Excel.Application()

    ↓
Open Workbook

    ↓
GetNamedRange() 取得命名範圍

    ↓
CloseWorkbook()

    ├─ _workbook.Close(false)
    ├─ _excelApp.Quit()
    ├─ Marshal.ReleaseComObject()
    ├─ GC.Collect()
    └─ GC.WaitForPendingFinalizers()

```
---
# 🧩IWordService 介面筆記
```csharp
Namespace： ExcelToWord_Service
Word 操作服務介面 — 定義所有 Word 相關操作的契約

此介面負責提供 Word 文件建立、貼圖、儲存與關閉 的標準操作，
與 IExcelService 搭配實現 Excel → Word 自動化報表流程

## 📘 介面原始程式碼
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// Word 操作服務介面
    /// 定義所有 Word 相關操作的契約
    /// </summary>
    public interface IWordService
    {
        /// <summary>
        /// 開啟或建立 Word 文件
        /// 類似 Excel 中的 _excelApp.Workbooks.Open(path);
        /// </summary>
        /// <param name="path">Word 檔案路徑</param>
        /// <returns>回傳 Word 文件物件 (Word.Document)</returns>
        Word.Document OpenOrCreate(string path);

        /// <summary>
        /// 插入 Excel 範圍的截圖到 Word
        /// </summary>
        /// <param name="doc">指定要貼入圖片的 Word 文件</param>
        /// <param name="sheetName">Excel 工作表名稱，用於插入小標題或標籤</param>
        /// <param name="range">Excel 範圍物件 (Range)，會被複製為圖片貼上</param>
        /// <param name="widthCm">貼入 Word 後的圖片寬度（單位：公分）</param>
        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm);

        /// <summary>
        /// 儲存並關閉 Word 文件，負責「儲存 + 關閉單一文件」
        /// </summary>
        /// <param name="doc">要儲存與關閉的 Word 文件</param>
        /// <param name="path">儲存的完整路徑</param>
        void SaveAndClose(Word.Document doc, string path);

        /// <summary>
        /// 關閉整個 Word 程式本體 (Application)
        /// </summary>
        void Quit();
    }
}