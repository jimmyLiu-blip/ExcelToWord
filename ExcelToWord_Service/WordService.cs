using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelToWord_Service
{
    /// Word 操作服務實作類別
    /// 負責建立 Word 文件、插入圖片、格式設定
    public class WordService : IWordService
    {
        private readonly Word.Application _wordApp;

        // Word是輸出端，要等資料準備好才建立/開始 <=> Excel是資料來源，一開始就要開啟資料夾
        // 在 WordService 的建構子裡，只要建立 Word 應用程式實例就好：
        // 在 Excel中，DisplayAlerts是布林值，所以可以使用 True / False
        // 在 Word中，DisplayAlerts是列舉型別，
        // 1. wdAlertsNone：不顯示任何提示
        // 2. wdAlertsMessageBox：只顯示關鍵錯誤提示
        // 3. wdAlertsAll：顯示所有提示（預設）
        // Word.WdAlertLevel 是 「警示層級列舉（Alert Level Enum）」
        public WordService()
        {
            _wordApp = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
        }

        public Word.Document OpenOrCreate(string path)
        {
            // 是一個三元運算子，條件 ? 條件為真時執行的程式 : 條件為假時執行的程式;
            // File.Exists(path) 會回傳 True / False
            return File.Exists(path)
                ? _wordApp.Documents.Open(path)
                : _wordApp.Documents.Add();
        }

        public void InsertRangePicture(Word.Document doc, string sheetName, Excel.Range range, float widthCm)
        {
            try
            {
                // 移到文件末端
                // doc.Content：回傳這份文件整個內容的 Range（從檔頭到檔尾的“選區”）
                // Collapse(direction)：把這個 Range 壓縮成長度 0 的“游標點”。
                // WdCollapseDirection：是 Word 內建的一個「列舉型別 (Enum)」
                // 用來指定 Range 折疊 (Collapse) 時要往哪個方向折疊。
                // wdCollapseStart：把 Range 的起訖端點都移到「起點」→ 游標停在內容最前方。
                // wdCollapseEnd：把 Range 的起訖端點都移到「終點」→ 游標停在內容最後方。
                doc.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // 插入標題
                // doc.Content.Paragraphs：文件中「所有段落」的集合。
                // .Add()：在 doc.Content 目前所在位置新增一個段落，並回傳新段落的 Paragraph 物件。
                // para：代表新建立的段落。
                // para.Range，這個 Range 就是「這個段落的文字範圍」。
                // para.Range.Text = $"【{sheetName}】";
                // 替這個段落的 Range 指定文字內容（會覆蓋原內容）。
                // 指派 Text 不會自動結尾換行，所以你後面才會 InsertParagraphAfter 加段落符號。
                // para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                // 設定該 Range 的內建樣式為「標題 2」（Heading 2）
                // set_Style(...) 用的是內建樣式列舉 WdBuiltinStyle，也可以改成自訂樣式名稱（字串）：
                // 好處：標題會自動套用字體大小、粗細、段前段後間距，日後也能用目錄產生
                // para.Range.InsertParagraphAfter();
                // 在 para.Range 的後面插入一個段落標記（相當於按一次 Enter）。
                // 作用：讓標題下方有一個「新行」，圖片就能貼在標題下方，不會擠在同一行。
                var para = doc.Content.Paragraphs.Add();
                para.Range.Text = $"【{sheetName}】";
                para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                para.Range.InsertParagraphAfter();

                // 複製 Excel 範圍為圖片
                // range.CopyPicture()：用來把指定範圍的儲存格內容「以圖片形式」複製到剪貼簿
                // Excel.XlPictureAppearance：是一個 列舉型別 (Enum)
                // 用來指定要以「螢幕呈現」還是「列印品質」的方式複製圖片
                // Excel.XlPictureAppearance.xlScreen：以螢幕顯示效果截圖（畫面上看到的顏色、陰影、邊框），✅ 通常用這個
                // Excel.XlPictureAppearance.xlPrinter：以列印輸出品質截圖（較高解析度、考慮列印樣式）
                // Excel.XlCopyPictureFormat：也是 列舉型別 (Enum)
                // 用來指定要以「圖片格式」或「向量圖格式」複製。
                // Excel.XlCopyPictureFormat.xlPicture：向量圖格式 (Metafile)，可縮放不失真，✅ 一般建議用這個
                // Excel.XlCopyPictureFormat.xlBitmap：點陣圖格式 (Bitmap)，固定解析度，放大會糊掉
                range.CopyPicture(
                    Excel.XlPictureAppearance.xlScreen,
                    Excel.XlCopyPictureFormat.xlPicture
                );

                // 啟用文件並貼上
                // doc 是一個 Word.Document 物件，（代表你目前開啟或新建的 Word 檔）。
                // doc.Activate()：讓這份 Word 文件成為 Word 應用程式目前的「作用中文件 (ActiveDocument)」。
                // _wordApp 是你的 Word.Application 物件。
                // Selection 屬性 = 「目前游標所在的範圍」。
                // 可以透過這個物件控制文字輸入、格式設定、圖片貼上等等。
                // .EndKey(Unit: Word.WdUnits.wdStory)：游標（Selection）移動到文件的「最後面」。
                // Word.WdUnits：列舉
                // wdCharacter：移動一個字元；wdLine：移動一行；wdParagraph：移動一個段落
                // wdPage：移動一頁；wdStory：移動到整份文件的開頭或結尾
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
                Console.WriteLine($"貼上圖片時發生問題：{ex.Message}");
                Console.ResetColor();
            }
        }

        /// 設定圖片大小（私有輔助方法）
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
                Console.WriteLine($"設定圖片大小失敗：{ex.Message}");
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