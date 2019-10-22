using System;
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
        [STAThread]
        static void Main(string[] args)
        {
            Console.Title = "Excel2PuompImportExcel v." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            try
            {               
                dynamic file = Dialog.FileBrowser();

                if (file.GetType() == typeof(string) ? file.Length > 0 : false)
                {
                    var excel = new ExcelProc(file);
                    excel.Run();
                } else if (file.GetType() == typeof(bool) ? true : false) {
                    Console.WriteLine($"Файл не выбран!");
                }
                Console.WriteLine($"Готово!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine($"\nнажмите любую клавишу...");
                Console.ReadKey();
            }            
        }
    }
}
