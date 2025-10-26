using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// Excel 操作服務介面
    /// 定義所有 Excel 相關操作的契約
    /// </summary>
    public interface IExcelService
    {
        /// <summary>取得工作簿</summary>
        Excel.Workbook Workbook { get; }

        /// <summary>取得指定工作表的命名範圍</summary>
        Excel.Range GetNamedRange(Excel.Worksheet ws, string rangeName);

        /// <summary>關閉 Excel 並釋放資源</summary>
        void Close();
    }
}