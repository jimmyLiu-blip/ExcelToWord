using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToWord.Service
{
    public interface IWordService
    {
        Word.Document OpenOrCreate(string wordpath);

        void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm);

        void SaveAndClose(Word.Document doc, string wordpath);

        void Close();
    }
}