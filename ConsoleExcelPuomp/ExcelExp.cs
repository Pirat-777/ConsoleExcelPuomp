using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleExcelPuomp
{
    class ExcelExp
    {
        public List<string[][]> Run(string file)
        {
            var excel = new ExcelProc(file);
            return this.Proc(excel);
        }

        private List<string[][]> Proc(ExcelProc pair)
        {
            string vidMedHelp, cnt = "0", summ = "0", mo = "", smo = "", moCode="";

            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "." };
                        
            List<string[][]> listMass = new List<string[][]>();
            int nn = 0;
            foreach (KeyValuePair<int, int> keyValue in pair.Run())
            {
                string dt = DateTime.Now.ToString("yyyyMMdd_hhmmss");
                int loop = keyValue.Value - (keyValue.Key + 2);
                var smokode = GetSmoKode(
                                        pair.xlSht.Range["E" + (3 + keyValue.Key)].Text +
                                        pair.xlSht.Range["G" + (3 + keyValue.Key)].Text);
                Console.WriteLine($"\nБлок № {++nn}:");
                foreach (KeyValuePair<string, string[]> item in smokode)
                {                    
                    smo = item.Key;
                    mo = pair.xlSht.Range["B" + keyValue.Key].Text.Trim();
                    moCode = new String(mo.Where(Char.IsDigit).ToArray()).Length >= 6 ? mo.Substring(0, 6).Trim():"";
                    
                    if (moCode.Length != 6)
                    {
                        pair.xlWB.Close();
                        throw new Exception($"\nОшибка: не указан или неверный код МО (см. ячейку {"B" + keyValue.Key}): \"{mo}\"");
                    }

                    Console.WriteLine($"Обработка MO: {mo}, SMO: {smo},");
                    string[][] newMass = new string[loop-5][];
                    for (int i = 6; i <= loop; i++)
                    {
                        vidMedHelp = pair.xlSht.Range["B" + (keyValue.Key + i)].Text.Trim();

                        foreach (var subItem in item.Value)
                        {
                            if ((new[] { "E", "G", }).Contains(subItem))
                                cnt = double.TryParse(pair.xlSht.Range[subItem + (keyValue.Key + i)].Text.Replace(" ", ""), out double res) ? Math.Round(res,2).ToString() : "0";

                            if ((new[] { "F", "H", }).Contains(subItem))
                                summ = double.TryParse(pair.xlSht.Range[subItem + (keyValue.Key + i)].Text.Replace(" ", ""), out double res) ? Math.Round(res,2).ToString() : "0";
                        }
                        //Console.WriteLine($"{dt}, МО:{moCode}, СМО:{smo}, Вид:{vidMedHelp},Кол-во:{cnt},Сумма:{summ}");
                        newMass[i-6] = new string[6] { dt, moCode, smo, vidMedHelp, cnt, summ };
                    }                    
                    listMass.Add(newMass);                    
                }                
            }            
            pair.xlWB.Close();            
            return listMass;
        }

        private Dictionary<string, string[]> GetSmoKode(string SMOName)
        {
            Dictionary<string, string[]> mass = new Dictionary<string, string[]>();

            if (SMOName.ToUpper().Contains("КАПИТАЛ"))
            {
                mass.Add("07004", new string[] { "E", "F" });
            }

            if (SMOName.ToUpper().Contains("РЕСО"))
            {
                mass.Add("07005", new string[] { "G", "H" });
            }
            return mass;
        }
    }
}
