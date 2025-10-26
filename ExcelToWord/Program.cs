using ExcelToWord_Configurement;
using ExcelToWord_Service;
using System;

namespace ExcelToWord
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("🚀 ExcelToWord 匯出系統啟動中...\n");

            try
            {
                // 載入設定
                ExportSettings settings = new ExportSettings();

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
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("❌ 發生錯誤：" + ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("\n按任意鍵結束...");
            Console.ReadKey();
        }
    }
}