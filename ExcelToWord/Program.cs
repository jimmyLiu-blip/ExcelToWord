using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelToWordByGroup_FinalFixed
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            string excelPath = @"C:\Reports\5GNR_3.7GHz_4.5GHz.xlsx";
            string outputFolder = @"C:\Reports\WordOutputs_ByItem";
            string[] targetNames = { "ACL_1", "ACLN_1" };
            int startSheetIndex = 7;

            Directory.CreateDirectory(outputFolder);

            Excel.Application excelApp = new Excel.Application();
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = false;

            var workbook = excelApp.Workbooks.Open(excelPath);

            try
            {
                for (int i = startSheetIndex; i <= workbook.Sheets.Count; i++)
                {
                    var ws = (Excel.Worksheet)workbook.Sheets[i];
                    Console.WriteLine($"\n▶ 處理工作表：{ws.Name}");

                    foreach (var rangeName in targetNames)
                    {
                        Excel.Range range = null;
                        try
                        {
                            range = workbook.Names.Item(rangeName).RefersToRange;
                        }
                        catch
                        {
                            try { range = ws.Names.Item(rangeName).RefersToRange; } catch { }
                        }

                        if (range == null)
                        {
                            Console.WriteLine($"⚠ 找不到命名範圍：{rangeName}（在 {ws.Name}）");
                            continue;
                        }

                        // Word 檔名
                        string itemName = rangeName.Contains("_") ? rangeName.Split('_')[0] : rangeName;
                        string wordPath = Path.Combine(outputFolder, $"{itemName}.docx");

                        // 【每次都開啟文件】
                        Word.Document doc;
                        if (File.Exists(wordPath))
                        {
                            doc = wordApp.Documents.Open(wordPath);
                        }
                        else
                        {
                            doc = wordApp.Documents.Add();
                        }

                        // 移到文件末端
                        doc.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        // 插入標題
                        var para = doc.Content.Paragraphs.Add();
                        para.Range.Text = $"【{ws.Name}】";
                        para.Range.set_Style(Word.WdBuiltinStyle.wdStyleHeading2);
                        para.Range.InsertParagraphAfter();

                        // 複製 Excel 範圍
                        range.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                          Excel.XlCopyPictureFormat.xlPicture);

                        // 啟用文件並移到末端貼上
                        doc.Activate();
                        wordApp.Selection.EndKey(Unit: Word.WdUnits.wdStory);
                        wordApp.Selection.Paste();
                        wordApp.Selection.TypeParagraph();

                        // 🟩 統一貼上圖片大小（寬度 15 cm）
                        foreach (Word.InlineShape shape in doc.InlineShapes)
                        {
                            try
                            {
                                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                                shape.Width = wordApp.CentimetersToPoints(15);  // 調整統一寬度
                            }
                            catch { }
                        }

                        // 加空行讓圖片之間保持距離
                        wordApp.Selection.TypeParagraph();
                        wordApp.Selection.TypeParagraph();

                        // 【立即存檔並關閉】
                        doc.SaveAs2(wordPath);
                        doc.Close();

                        Console.WriteLine($"✅ 匯出 {rangeName} → {wordPath}");

                        // 釋放文件物件
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);

                        // 短暫延遲防止 COM 阻塞
                        System.Threading.Thread.Sleep(150);
                    }
                }

                Console.WriteLine("\n🎉 全部完成！");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"❌ 發生錯誤：{ex.Message}");
                Console.WriteLine($"詳細資訊：{ex.StackTrace}");
                Console.ResetColor();
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
                wordApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
