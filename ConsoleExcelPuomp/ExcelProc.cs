using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleExcelPuomp
{
    class ExcelProc
    {
        private Excel.Application xlApp; //Excel
        public Excel.Workbook xlWB; //рабочая книга            
        public Excel.Worksheet xlSht; //лист Excel            
        public Excel.Range Rng; //диапазон ячеек  
        private string file;

        public ExcelProc(string file, int WorksheetNum = 1)
        {
            this.file = file;
            this.xlApp = new Excel.Application();
            this.xlWB = xlApp.Workbooks.Open(file);

            //название листа xlWB.Worksheets["Лист1"] или 1-й лист в книге xlWB.Worksheets[1];
            this.xlSht = xlWB.Worksheets[WorksheetNum];
        }

        public Dictionary<int, int> Run()
        {
            return GetStartEndPair();            
        }

        private Dictionary<int, int> GetStartEndPair()
        {
            Dictionary<int,int> massStartEndPair = new Dictionary<int,int>();
            int j = 0, s = 0, e = 0;
            for (int i = 1; i <= GetLastRow(); i++)
            {

                string x = xlSht.Range["A" + i].Text;
                if (x.Length > 0)
                {
                    if (x.Trim().ToUpper()[0].ToString() == "Н")
                    {
                        s = i;
                        j++;
                    }
                        
                    if (x.Trim().ToUpper()[0].ToString() == "К")
                    {
                        e = i;
                        j++;
                    }     
                    
                    if (j == 2)
                    {
                        massStartEndPair.Add(s, i);
                        j = 0; s = 0; e = 0;
                    }
                }
            }
            //вызов ошибки при непарном начале и конце
            if (j % 2 != 0)
                throw new Exception($"непарные начало или конец в столбце \"А\" документа \"{Path.GetFileName(file)}\"");

            return massStartEndPair;
        }

        //последняя заполненная строка в столбце А
        private int GetLastRow(string ColumnTitle = "A")
        {
            return xlSht.Cells[xlSht.Rows.Count, ColumnTitle].End[Excel.XlDirection.xlUp].Row;  
        }

        //последний заполненный столбец в 1-й строке
        private int GetLastColumn(int RowNum = 1)
        {
            return xlSht.Cells[RowNum, xlSht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; 
        }
    }
}
