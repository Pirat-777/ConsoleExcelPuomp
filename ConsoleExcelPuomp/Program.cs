using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConsoleExcelPuomp
{
    class Program
    {
        public static string protocolDate;
        public static string protocolNum;

        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "Excel2PuompImportExcel v." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            try
            {
                dynamic file = Dialog.FileBrowser();

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
