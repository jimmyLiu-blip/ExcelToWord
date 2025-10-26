using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord_Service
{
    /// <summary>
    /// Word 操作服務實作類別
    /// 負責建立 Word 文件、插入圖片、格式設定
    /// </summary>
    public class WordService : IWordService
    {
        private readonly Word.Application _wordApp;

        public WordService()
        {
            _wordApp = new Word.Application
            {
                Visible = false
            };
        }

        public Word.Document OpenOrCreate(string path)
        {
            return System.IO.File.Exists(path)
                ? _wordApp.Documents.Open(path)
                : _wordApp.Documents.Add();
        }

        public void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm)
        {
            try
            {
                // 移到文件末端
                doc.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // 插入標題
                var para = doc.Content.Paragraphs.Add();
                para.Range.Text = $"【{sheetName}】";
                para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                para.Range.InsertParagraphAfter();

                // 複製 Excel 範圍為圖片
                range.CopyPicture(
                    Excel.XlPictureAppearance.xlScreen,
                    Excel.XlCopyPictureFormat.xlPicture
                );

                // 啟用文件並貼上
                doc.Activate();
                _wordApp.Selection.EndKey(Unit: Word.WdUnits.wdStory);
                _wordApp.Selection.Paste();

                // 【關鍵修正】統一圖片大小 - 使用文件的 InlineShapes 集合
                SetImageSize(doc, widthCm);

                _wordApp.Selection.TypeParagraph();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"⚠ 貼上圖片時發生問題：{ex.Message}");
                Console.ResetColor();
            }
        }

        /// <summary>設定圖片大小（私有輔助方法）</summary>
        private void SetImageSize(Word.Document doc, float widthCm)
        {
            try
            {
                // 【方法1：取得文件中最後一張圖片（剛貼上的）】
                if (doc.InlineShapes.Count > 0)
                {
                    // 取得最後一張圖片
                    var shape = doc.InlineShapes[doc.InlineShapes.Count];
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Width = _wordApp.CentimetersToPoints(widthCm);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠ 設定圖片大小失敗：{ex.Message}");
            }
        }

        public void SaveAndClose(Word.Document doc, string path)
        {
            doc.SaveAs2(path);
            doc.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
        }

        public void Quit()
        {
            _wordApp?.Quit();
            if (_wordApp != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_wordApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}