using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleExcelPuomp
{
    class Import2Excel
    {
        private List<string[][]> res;        
        
        public Import2Excel(List<string[][]> res)
        {
            this.res = res;                        
        }

        public void Run()
        {
            // Создаём экземпляр нашего приложения
            Excel.Application excelApp = new Excel.Application();
            // Создаём экземпляр рабочий книги Excel
            Excel.Workbook workBook;
            // Создаём экземпляр листа Excel
            Excel.Worksheet workSheet;

            Directory.CreateDirectory(Environment.CurrentDirectory+"\\Результат");
            string pathSave = Environment.CurrentDirectory + "\\Результат\\";

            string file = "";
            foreach (var item in res)
            {     
                workBook = excelApp.Workbooks.Add();
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
                //Console.WriteLine($"{dt}, МО:{moCode}, СМО:{smo}, Вид:{vidMedHelp},Кол-во:{cnt},Сумма:{summ}");
                
                workSheet.Cells[1, 1] = "признак";
                workSheet.Cells[1, 2] = "код по справочнику";
                workSheet.Cells[1, 3] = "виды медицинской помощи";
                workSheet.Cells[1, 4] = "реабилитация";
                workSheet.Cells[1, 5] = "эко";
                workSheet.Cells[1, 6] = "онкология";
                workSheet.Cells[1, 7] = "объём";
                workSheet.Cells[1, 8] = "сумма";
                workSheet.Cells[1, 9] = "мо";
                workSheet.Cells[1, 10] = "смо";
                workSheet.Cells[1, 11] = "ДАТА ДОК";
                workSheet.Cells[1, 12] = "ПРОТОКОЛ";

                for (int i = 1; i <= item.Length; i++)
                {
                    Console.WriteLine($"{item[i - 1][1]}, {item[i - 1][2]}, {item[i - 1][3]}, {item[i - 1][4]}, {item[i - 1][5]}");

                    file = item[i - 1][1] + "_" + item[i - 1][2] + "_" + item[i - 1][0] + ".xls";                   

                    workSheet.Cells[i + 1, 1] = GetVidMH(item[i - 1][3])[0];
                    workSheet.Cells[i + 1, 2] = GetVidMH(item[i - 1][3])[1];
                    workSheet.Cells[i + 1, 3] = item[i - 1][3];
                    workSheet.Cells[i + 1, 4] = GetVidMH(item[i - 1][3])[2];
                    workSheet.Cells[i + 1, 5] = GetVidMH(item[i - 1][3])[3];
                    workSheet.Cells[i + 1, 6] = GetVidMH(item[i - 1][3])[4];
                    workSheet.Cells[i + 1, 7] = item[i - 1][4];
                    workSheet.Cells[i + 1, 8] = item[i - 1][5];
                    workSheet.Cells[i + 1, 9] = "\'" + item[i - 1][1];
                    workSheet.Cells[i + 1, 10] = "\'" + item[i - 1][2];
                    workSheet.Cells[i + 1, 11] = Program.protocolDate;
                    workSheet.Cells[i + 1, 12] = Program.protocolNum;                    
                }

                Console.WriteLine($"\nСохранение {file} в {pathSave}");
                workBook.SaveAs($"{pathSave+file}", Excel.XlFileFormat.xlExcel8);
                workBook.Close();
                excelApp.Quit();
            }            
        }

        private string[] GetVidMH(string str)
        {            
            if (GetClearStr(str).Contains("СТАЦИОНАРКСГ"))
                return new string[] { "КССВОД", "", "", "", "" };
            if (GetClearStr(str).Contains("ОНКОЛОГИЯСТАЦИОНАР"))
                return new string[] { "КССВОД", "", "0", "0", "1" };
            if (GetClearStr(str).Contains("РЕАБИЛИТАЦИЯЗСЛ"))
                return new string[] { "КССВОД", "", "1", "0", "0" };
            if (GetClearStr(str).Contains("ВМПЗСЛ"))
                return new string[] { "ВМП", "99", "", "", "" };
            if (GetClearStr(str).Contains("ДНЕВНОЙСТАЦИОНАРКСГ"))
                return new string[] { "ДССВОД", "", "", "", "" };
            if (GetClearStr(str).Contains("ОНКОЛОГИЯДНСТАЦИОНАР"))
                return new string[] { "ДССВОД", "", "0", "0", "1" };
            if (GetClearStr(str).Contains("ДИАЛИЗВУСЛДНСТАЦИОНАРАУСЛУГА"))
                return new string[] { "АМБ", "24", "", "", "" };
            if (GetClearStr(str).Contains("ЭКО"))
                return new string[] { "ДССВОД", "", "0", "1", "0" };
            if (GetClearStr(str).Contains("НЕОТЛОЖНЫЕПОСЕЩЕНИЯ"))
                return new string[] { "АМБ", "9", "", "", "" };
            if (GetClearStr(str).Contains("ОБРАЩЕНИЯ"))
                return new string[] { "АМБ", "11", "", "", "" };
            if (GetClearStr(str).Contains("ПРОФИЛАКТИЧЕСКИЕПОСЕЩЕНИЯРАЗОВЫЕ"))
                return new string[] { "АМБ", "8", "", "", "" };
            if (GetClearStr(str).Contains("ПОДУШЕВОЕФИНАНСИРОВАНИЕ"))
                return new string[] { "АМБ", "35", "", "", "" };
            if (GetClearStr(str).Contains("ДИАГНОСТИЧЕСКИЕПОСЕЩЕНИЯ"))
                return new string[] { "АМБ", "34", "", "", "" };
            if (GetClearStr(str).Contains("ДИАЛИЗВУСЛАПП"))
                return new string[] { "АМБ", "23", "", "", "" };
            if (GetClearStr(str).Contains("СТОМАТОЛОГИЧЕСКАЯПОМОЩЬУЕТ"))
                return new string[] { "АМБ", "26", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯДЕТЕЙСИРОТЗСЛ"))
                return new string[] { "АМБ", "1", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯВЗРОСЛЫХ1ЫЙЭТАПЗСЛ"))
                return new string[] { "АМБ", "6", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯВЗРОСЛЫХ2ОЙЭТАПЗСЛ"))
                return new string[] { "АМБ", "7", "", "", "" };
            if (GetClearStr(str).Contains("ПРОФОСМОТРВЗРОСЛЫХЗСЛ"))
                return new string[] { "АМБ", "8", "", "", "" };
            if (GetClearStr(str).Contains("МЕДОСМОТРНЕСОВЕРШЕННОЛЕТНИХ1ЭТЗСЛ"))
                return new string[] { "АМБ", "3", "", "", "" };
            if (GetClearStr(str).Contains("МЕДОСМОТРНЕСОВЕРШЕННОЛЕТНИХ2ЭТЗСЛ"))
                return new string[] { "АМБ", "28", "", "", "" };
            if (GetClearStr(str).Contains("СКОРАЯСПЕЦМЕДПОМОЩЬПОДУШНОР"))
                return new string[] { "СМП", "16", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯВЗРОСЛЫХ1ЫЙЭТЗСЛ13"))
                return new string[] { "АМБ", "6", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯВЗРОСЛЫХ1ЫЙЭТЗСЛ12"))
                return new string[] { "АМБ", "27", "", "", "" };
            if (GetClearStr(str).Contains("ДИСПАНСЕРИЗАЦИЯВЗРОСЛЫХ2ОЙЭТЗСЛ"))
                return new string[] { "АМБ", "7", "", "", "" };            
            if (GetClearStr(str).Contains("СКОРАЯСПЕЦМЕДПОМОЩЬПОДУШНОР"))
                return new string[] { "СМП", "16", "", "", "" };

            return new string[] { "","","","","","","" };
        }
        private string GetClearStr(string str)
        {
            foreach (var item in new string[] {" ","(",")","\\","/","-",":",".","," })
            {
                str = str.Replace(item,"");
            }
            //Console.WriteLine(str);
            return str.ToUpper();
        }
    }
}
