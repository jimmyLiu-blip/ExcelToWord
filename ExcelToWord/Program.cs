using ExcelToWord.Configuration;
using ExcelToWord.Service;
using ExcelToWord_Service;
using System;

namespace ExcelToWord
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("=================================");
            Console.WriteLine(" ExcelToWord 匯出系統啟動中...");
            Console.WriteLine("=================================\n");

            try
            {
                // 載入設定
                ExportSettings settings = new ExportSettings();

                // 顯示設定資訊
                Console.WriteLine($"Excel 檔案: {settings.ExcelPath}");
                Console.WriteLine($"輸出資料夾: {settings.OutputFolder}");
                Console.WriteLine($"目標範圍: {string.Join(", ", settings.TargetNames)}");
                Console.WriteLine($"起始工作表: 第 {settings.StartSheetIndex} 張\n");

                // 建立服務實例
                IExcelService excelService = new ExcelService(settings.ExcelPath);
                IWordService wordService = new WordService();

                // 建立協調器並執行
                ExportCoordinator coordinator = new ExportCoordinator(
                    settings,
                    excelService,
                    wordService
                );

                coordinator.Run();

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n所有作業已完成!");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n發生錯誤：{ex.Message}");
                Console.WriteLine($"錯誤詳情：{ex.StackTrace}");
                Console.ResetColor();
            }

            Console.WriteLine("\n按任意鍵結束...");
            Console.ReadKey();
        }
    }
}