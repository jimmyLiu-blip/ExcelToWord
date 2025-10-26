using Excel = Microsoft.Office.Interop.Excel;
// 的意思是：「幫 Microsoft.Office.Interop.Excel 這個很長的命名空間取一個簡短的名字叫 Excel。」
// 也可以直接寫 using Microsoft.Office.Interop.Excel; 但是Word也會有類似用法，避免衝突產生，會建議取別名

namespace ExcelToWord_Service
{
    /// Excel 操作服務介面
    /// 定義所有 Excel 相關操作的契約
    public interface IExcelService
    {
        /// 取得工作簿
        /// 這是一個唯獨屬性，提供「目前開啟的 Excel 活頁簿」給外部使用。
        /// Excel.Workbook：屬性的型別 → 代表一個 Excel 檔案（活頁簿）
        /// Workbook：屬性的名稱
        Excel.Workbook Workbook { get; }

        /// 取得指定工作表的命名範圍
        /// Excel.Range：方法的「回傳型別」 → 表示會回傳一個範圍物件（Range）
        /// 型別：Excel.Worksheet，含義：Excel 裡的一張「工作表」，名稱：ws（只是變數名，常見縮寫 = worksheet）
        /// 這裡的 string rangeName 代表 名稱管理員 ➜ 定義名稱「ACL_1」指向一塊儲存格範圍
        Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName);

        /// 關閉 Excel 並釋放資源
        void Close();
    }
}

// Excel.Application → Excel 程式本體，代表 整個 Excel 應用程式本身。
// 你開啟一個 Excel.exe，其實就是在建立一個 Excel.Application 物件。

// Excel.Workbook → 活頁簿 📓（Excel 檔案）
// 代表一個「Excel 檔案」（例如 5GNR_3.7GHz_4.5GHz.xlsx）。
// 一個 Excel 應用程式 (Application) 可以同時開多個 Workbook。

// Excel.Worksheet → 工作表 📄（檔案中的分頁）
// 代表一張「Excel 工作表」，就是你在檔案下方看到的「Sheet1」「Sheet2」「n40_10M」這些分頁。

// Excel.Range → 儲存格或範圍
// 代表一格或多格儲存格的範圍。