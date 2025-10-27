using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// Word 操作服務介面
    /// 定義所有 Word 相關操作的契約
    public interface IWordService
    {
        /// 開啟或建立 Word 文件
        /// 類似Excel中的_excelApp.Workbooks.Open(path);
        Word.Document OpenOrCreate(string path);

        /// 插入 Excel 範圍的截圖到 Word
        /// 
        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm);

        /// 儲存並關閉 Word 文件，負責「儲存 + 關閉單一文件」
        void SaveAndClose(Word.Document doc, string path);

        /// 負責「關閉整個 Word 程式本體」
        void Quit();
    }
}
