using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.CodeDom;

namespace ExcelToWord.Service
{
    public class WordService : IWordService
    {
        private readonly Word.Application _wordApp;

        public WordService()
        {
            _wordApp = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
        }

        public Word.Document OpenOrCreate(string wordPath)
        {
            return File.Exists(wordPath)
                ? _wordApp.Documents.Open(wordPath)
                : _wordApp.Documents.Add();
        }

        public void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm)
        {
            try
            {
                doc.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                var para = doc.Content.Paragraphs.Add();
                para.Range.Text = $"【{sheetName}】";
                para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2); //多了Unit
                para.Range.InsertParagraphAfter(); //整行忘記

                range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture); //整行忘記 + 順序有誤

                doc.Activate();

                _wordApp.Selection.EndKey(Unit: Word.WdUnits.wdStory); //整行忘記，是Selection不是Section，前面跟著Applicaiotn
                _wordApp.Selection.Paste(); //整行忘記，是Selection不是Section，前面跟著Applicaiotn

                SetImageSize(doc, widthCm);

                _wordApp.Selection.TypeParagraph(); //整行忘記，是Selection不是Section，前面跟著Applicaiotn
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"出現異常錯誤無法貼上圖片，{ex.Message}");
                Console.ResetColor();
            }

        }

        private void SetImageSize(Word.Document doc, float widthCm)
        {
            try
            {
                if (doc.InlineShapes.Count > 0)
                {
                    var shape = doc.InlineShapes[doc.InlineShapes.Count];
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Width = _wordApp.CentimetersToPoints(widthCm); // CentimentersToPoints拼錯，前面跟著Applicaiotn
                }
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"出現異常錯誤無法調整Word圖片大小，{ex.Message}");
                Console.ResetColor();
            }
        }

        public void SaveAndClose(Word.Document doc, string wordPath)
        {
            try
            {
                doc?.SaveAs2(wordPath);
                doc?.Close(false);

                if (doc != null)
                {
                    Marshal.FinalReleaseComObject(doc);
                }

                doc = null;
            }

            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"發生異常錯誤，無法關閉Word檔案，{ex.Message}");
                Console.ResetColor();
            }
        }

        public void Close()
        {
            try
            {
                _wordApp?.Quit();

                if (_wordApp != null)
                {
                    Marshal.FinalReleaseComObject(_wordApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                Console.BackgroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"發生異常錯誤，無法關閉Word執行檔，{ex.Message}");
                Console.ResetColor();
            }
        }
    }
}