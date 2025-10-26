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
        /// <summary>開啟或建立 Word 文件</summary>
        Word.Document OpenOrCreate(string path);

        /// <summary>插入 Excel 範圍的截圖到 Word</summary>
        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm);

        /// <summary>儲存並關閉 Word 文件</summary>
        void SaveAndClose(Word.Document doc, string path);

        /// <summary>關閉 Word 應用程式</summary>
        void Quit();
    }
}
