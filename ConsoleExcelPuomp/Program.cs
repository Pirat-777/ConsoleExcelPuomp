using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;

namespace ConsoleExcelPuomp
{
    class SettingJson 
    {
        public string[] el { get; set; }
    }
        
    class Program
    {
        public static string protocolDate;
        public static string protocolNum;
        public static string settings= @"settings.json";
        public static Dictionary<string, string[]> settingsData;

        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "Excel2PuompImportExcel v." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            try
            {
                dynamic file = Dialog.FileBrowser();

                if(!File.Exists(settings))
                     throw new Exception( $"Файл настроект \"{settings}\" не найден!" );

                // чтение данных json                
                settingsData = JsonConvert.DeserializeObject<Dictionary<string, string[]>>(File.ReadAllText(settings, Encoding.Default));

                
                if (file.GetType() == typeof(string) ? file.Length > 0 : false)
                {
                    Console.Write($"Введите номер протокола: ");
                    protocolNum = Console.ReadLine();
                    Console.Write($"Введите дату протокола: ");
                    protocolDate = Console.ReadLine();

                    new Import2Excel(new ExcelExp().Run(file)).Run();
                }
                else if (file.GetType() == typeof(bool) ? true : false)
                {
                    Console.WriteLine($"Файл не выбран!");
                }
                Console.WriteLine($"\nРезультат: выполнено.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine($"\nДля закрытия нажмите любую клавишу...");
                Console.ReadKey();
            }            
        }        
    }
}
