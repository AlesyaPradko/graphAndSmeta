using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public static class ParserExcel
    {
       private readonly static object _misValue = System.Reflection.Missing.Value;
        //возвращает словарь, где ключ - номер по смете, значение - объем работ по позиции
        public static Dictionary<int, double> GetScopeWorkAktKSone(AktKS akt)
        {
            Dictionary<int, double> total = new Dictionary<int, double>();
            int valueNumPoz;
            double valueScopeWork;
            int count = 0;
            for (int j = akt.KeyNumberPosKS.Row + 4; j < akt.RangeDoc.Rows.Count; j++)
            {
                Excel.Range cellsNumPozColumnTabl = akt.SheetDoc.Cells[j, akt.KeyNumberPosKS.Column];
                Excel.Range cellsScopeColumnTabl = akt.SheetDoc.Cells[j, akt.KeyScopeWorkinAktKS.Column];
                if (cellsNumPozColumnTabl.Value2 == null)
                    count++;
                if (count > 10) break;
                if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsScopeColumnTabl != null && cellsScopeColumnTabl.Value2 != null && cellsScopeColumnTabl.Value2.ToString() != "" && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells && !cellsScopeColumnTabl.MergeCells)
                {
                    try
                    {
                        valueNumPoz = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                        valueScopeWork = Convert.ToDouble(cellsScopeColumnTabl.Value2);
                        total.Add(valueNumPoz, valueScopeWork);
                        count = 0;
                    }
                    catch (FormatException ex)
                    {
                        akt.Error += $" Вы ввели неверный формат для {akt.AddressDoc} в строке {cellsNumPozColumnTabl.Row} в столбце {cellsNumPozColumnTabl.Column} не должно быть [.,букв], только целые числа или же в столбце {cellsScopeColumnTabl.Column} не должно быть [.букв], только дробные числа с [,] или целые числа)\n";
                    }
                    catch (ArgumentException ex)
                    {
                        akt.Error += $"{ex.Message} Проверьте чтобы в {akt.AddressDoc} не повторялись значения позиций по смете в строке {cellsNumPozColumnTabl.Row}\n";
                    }
                }
            }
            return total;
        }
        //возвращает лист смет, находящихся в папке
        public static List<Smeta> GetAllSmeta(string userKS) 
        {
            string[] nameAdresSmeta = Directory.GetFiles(userKS);
            List<Smeta> containFolderSmeta = new List<Smeta>();
            Parallel.ForEach(nameAdresSmeta, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount / 2 }, oneNameKS =>
             {
                 containFolderSmeta.Add((Smeta)GetDoc(oneNameKS, false));
             });
            return containFolderSmeta;
        }
        //возвращает лист актов КС-2, находящихся в папке
        public static List<AktKS> GetAllAktKS(string userKS)
        {
            string[] nameAdresAktKS = Directory.GetFiles(userKS);
            List<AktKS> containFolderAkt = new List<AktKS>();
            Parallel.ForEach(nameAdresAktKS, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount / 2 }, oneNameKS =>
            {
                containFolderAkt.Add((AktKS)GetDoc(oneNameKS, true));
            });
            return containFolderAkt;
        }
        //возвращает акт КС-2 или смету
        public static DocumentExcel GetDoc(string oneNameKS, bool flag)
        {
            DocumentExcel doc= null;
            if (!oneNameKS.Contains("~$") && oneNameKS.Contains(".xlsx"))
            { 
                if(flag)
                    doc = new AktKS(oneNameKS);
                else
                    doc= new Smeta(oneNameKS); 
            }
            return doc;
        }
        //возвращает словарь, где ключ - название сметы, значение - лист с названием актов КС-2
        public static Dictionary<string, List<string>> GetContainAktKSinOneSmeta(List<AktKS> folderKS, List<Smeta> smeta)
        {
            int foundS1;
            Dictionary<string, List<string>> aktAllKSforOneSmeta = new Dictionary<string, List<string>>();          
            for (int u = 0; u < smeta.Count; u++)
            {
                string numberSmeta;
                MatchCollection mathesNumerSmeta = RegexReg.NameSmeta.Matches(smeta[u].AddressDoc);
                if (mathesNumerSmeta.Count > 0)
                {
                    numberSmeta = NameSmetaNumber(mathesNumerSmeta, out foundS1);
                    if (foundS1 != -1)
                    {
                        numberSmeta = numberSmeta.Remove(0, 1 + foundS1);
                        List<string> aktKSforSmeta = TakeListAktKS(folderKS, numberSmeta);
                        if (aktKSforSmeta.Count == 0)
                         smeta[u].Error += $"В актах КС отсутствует номер сметы или неверно записан, либо для сметы {smeta[u].AddressDoc} нет актов КС-2 \n";
                        aktAllKSforOneSmeta.Add(smeta[u].AddressDoc, aktKSforSmeta);
                    }
                }
            }
            return aktAllKSforOneSmeta;
        }
        //возвращает лист актов КС-2
        private static List<string> TakeListAktKS(List<AktKS> folderKS,string numberSmeta)
        {
            List<string> aktKSforSmeta = new List<string>();
            for (int c = 0; c < folderKS.Count; c++)
            {
                Excel.Range rangAktKS = folderKS[c].SheetDoc.get_Range("A1", "Q40");
                if (rangAktKS.Find(numberSmeta) == null) continue;
                else
                {
                    aktKSforSmeta.Add(folderKS[c].AddressDoc);
                }
            }
            return aktKSforSmeta;
        }
        //возвращает номер светы
        private static string NameSmetaNumber(MatchCollection mathesNumerSmeta,out int foundS1)
        {
            string numberSmeta = null;
            foreach (Match numSmeta in mathesNumerSmeta)
            {
                numberSmeta = numSmeta.Value;
            }
            foundS1 = numberSmeta.IndexOf("№");
            return numberSmeta;
        }
        //возвращает копию сметы, которую преобразуют в ведомость
        public static Smeta CopyExcelSmetaOne(Smeta smeta, string testUserWhereSave)
        {
            string noCopy = null;
            if (!File.Exists(testUserWhereSave))
            { 
              smeta.DocCur.SaveCopyAs(testUserWhereSave);
              noCopy = testUserWhereSave;
            }
            else
            {
                DialogResult result = MessageBox.Show($"Вы хотите заменить копию ведомости {testUserWhereSave}?", "Предупреждение", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    smeta.DocCur.SaveCopyAs(testUserWhereSave);
                    noCopy = testUserWhereSave;
                }
            }
            CloseDoc(smeta);
            Smeta smetaNew = null;
            if(noCopy != null)
            {
                smetaNew = new Smeta(testUserWhereSave);
            }      
            return smetaNew;
        }
       //закрывает документ иксель
        public static void CloseDoc(DocumentExcel doc)
        {
            doc.DocCur.Close(true, _misValue, _misValue);
        }
        //закрывает лист документов иксель
        public static void CloseAktKS(List<AktKS> listAktKStoOneSmeta)
        {
            for (int numKS = 0; numKS < listAktKStoOneSmeta.Count; numKS++)
            {
                listAktKStoOneSmeta[numKS].DocCur.Close(false, _misValue, _misValue);
            }
        }
        //метод удаляет ненужные строки и столбцы
        public static void DeleteColumnandRow(Smeta smeta, ref int lastRowCellsafterDelete)
        {
            try
            {
                DeleteRow(smeta, ref lastRowCellsafterDelete);
                DeleteColumn(smeta);
            }
            catch (ArgumentOutOfRangeException exc)
            {
                smeta.Error += $"{exc.Message} вы пытаетесь повторно удалить уже удаленные ячейки\n";
            }
        }
        //метод удаляет ненужные строки
        private static void DeleteRow(Smeta smeta, ref int lastRowCellsafterDelete)
        {
            int amountRow = 0;
            List<int> deleteExcessCells = new List<int>();
            for (int u = smeta.KeyNumberPosSmeta.Row + 5; u <= smeta.RangeDoc.Rows.Count; u++)
            {
                Excel.Range cellsFirstColumnTabl = smeta.SheetDoc.Cells[u, smeta.KeyNumberPosSmeta.Column];
                if (cellsFirstColumnTabl.MergeCells && !cellsFirstColumnTabl.Value.ToString().Contains("Раздел"))
                {
                    deleteExcessCells.Add(cellsFirstColumnTabl.Row);
                }
                if (cellsFirstColumnTabl.Value != null && cellsFirstColumnTabl.Value.ToString() != "")
                {
                    amountRow++;
                }
            }
            deleteExcessCells.Reverse();
            if (deleteExcessCells.Count != 0)
            {
                lastRowCellsafterDelete = deleteExcessCells[0] - deleteExcessCells.Count;
            }
            else
            {
                lastRowCellsafterDelete = smeta.KeyNumberPosSmeta.Row + 5 + amountRow; ///!!!!!!!
            }
            for (int u = smeta.RangeDoc.Rows.Count; u > smeta.KeyNumberPosSmeta.Row + 4; u--)
            {
                Excel.Range cellsFirstColumnTabl = smeta.SheetDoc.Cells[u, smeta.KeyNumberPosSmeta.Column];
                for (int v = 0; v < deleteExcessCells.Count; v++)
                {
                    if (cellsFirstColumnTabl.Row == deleteExcessCells[v])
                    {
                        Excel.Range lastColumnOnDelet = smeta.SheetDoc.Cells[cellsFirstColumnTabl.Row, smeta.RangeDoc.Columns.Count];
                        Excel.Range rowOnDelet = smeta.SheetDoc.get_Range(cellsFirstColumnTabl, lastColumnOnDelet);
                        rowOnDelet.Delete();
                        break;
                    }
                }
            }
        }
        //метод удаляет ненужные столбцы
        private static void DeleteColumn(Smeta smeta)
        {
            Regex rex = new Regex(@"(С|с)тоимость");
            MatchCollection mathesStoim = null;
            Excel.Range lastCellOnRangeForDelet = smeta.SheetDoc.Cells[smeta.RangeDoc.Rows.Count, smeta.RangeDoc.Columns.Count];
            Excel.Range firstCellOnRangeForDelet = null;
            for (int u = smeta.KeyNumberPosSmeta.Column; u <= smeta.RangeDoc.Columns.Count; u++)
            {
                Excel.Range cellsFirstRowTabl = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, u];
                if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value != null)
                {
                    mathesStoim = rex.Matches(cellsFirstRowTabl.Value.ToString());
                }
                if (mathesStoim.Count > 0)
                {
                    firstCellOnRangeForDelet = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, u];
                    break;
                }
            }
            if (firstCellOnRangeForDelet != null)
            {
                Excel.Range rangeOnDelet = smeta.SheetDoc.get_Range(firstCellOnRangeForDelet, lastCellOnRangeForDelet);
                rangeOnDelet.Delete();
            }
            else
            {
                string er = $" Проверьте чтобы в {smeta.AddressDoc} было верно записано устойчивое выражение [(С|с)тоимость]\n";
                smeta.Error += er;
                throw new NullValueException(er);
            }
        }
        //возвращает словарь, где ключ - номер по смете, значение - дефолтное значение
        public static Dictionary<int, T> GetkeySmetaForRecord<T>(Smeta smeta)
        {
            Dictionary<int, T> resultwithNumPoz = new Dictionary<int, T>();
            int numPozSmeta;
            T zerovalue = default(T);
            for (int j = smeta.RangeDoc.Row + 4; j <= smeta.RangeDoc.Count; j++)
            {
                Excel.Range cellsFirstColumnTabl = smeta.SheetDoc.Cells[j, smeta.RangeDoc.Column];
                if (cellsFirstColumnTabl != null && cellsFirstColumnTabl.Value2 != null && cellsFirstColumnTabl.Value2.ToString() != "" && !cellsFirstColumnTabl.MergeCells)
                {
                    try
                    {
                        numPozSmeta = Convert.ToInt32(cellsFirstColumnTabl.Value2);
                        zerovalue = default(T);
                        resultwithNumPoz.Add(numPozSmeta, zerovalue);
                    }
                    catch (ArgumentException ex)
                    {
                        if (zerovalue is double)
                            smeta.Error += $"{ex.Message} Проверьте чтобы в {smeta.AddressDoc} не повторялись значения позиций по смете в строке {cellsFirstColumnTabl.Row}\n";
                    }
                    catch (FormatException ex)
                    {
                        if (zerovalue is double)
                            smeta.Error += $"{ex.Message} Вы ввели неверный формат для {smeta.AddressDoc} в строке {cellsFirstColumnTabl.Row} в столбце {cellsFirstColumnTabl.Column}(не должно быть [., букв], только целые числа)\n";
                    }
                }
            }
            return resultwithNumPoz;
        }
        //возвращает дату акта КС-2
        public static string FindDateAktKS(Regex monthorYear, string find)
        {
            string dateMonthOrYear = null;
            MatchCollection yearMonth = monthorYear.Matches(find);
            if (yearMonth.Count > 0)
            {
                foreach (Match oneDate in yearMonth)
                {
                    dateMonthOrYear = oneDate.Value;
                }
            }
            return dateMonthOrYear;
        }
        //возвращает название месяца
        public static string MonthLetterInt(int montStart)
        {
            Console.WriteLine("MonthLetterInt");
            string monthLetter = null;
            switch (montStart)
            {
                case 1: monthLetter = "январь"; break;
                case 2: monthLetter = "февраль"; break;
                case 3: monthLetter = "март"; break;
                case 4: monthLetter = "апрель"; break;
                case 5: monthLetter = "май"; break;
                case 6: monthLetter = "июнь"; break;
                case 7: monthLetter = "июль"; break;
                case 8: monthLetter = "август"; break;
                case 9: monthLetter = "сентябрь"; break;
                case 10: monthLetter = "октябрь"; break;
                case 11: monthLetter = "ноябрь"; break;
                case 12: monthLetter = "декабрь"; break;
            }
            return monthLetter;
        }
        //возвращает полное название сметы
        public static string GetNameAktKS(AktKS akt)
        {
            string result = "Акт КС-2 №";
            string yearAktKS = akt.YearAktKS.ToString();
            string monthAktKSpropis = MonthLetterInt(akt.MonthAktKS);
            result += $"{akt.NumAktKS} {monthAktKSpropis} {yearAktKS}\n";
            return result;
        }
        //возвращает ячейку где фигурирует слово Количество
        public static Excel.Range FindCellOfRegulScope(Excel.Worksheet akt,Excel.Range ran)
        {
            MatchCollection mathesFindCell;
            Excel.Range findCellsColumn = null;
            for (int u = 1; u <= ran.Rows.Count; u++)
            {
                for (int j = 1; j <= ran.Columns.Count; j++)
                {
                    Excel.Range nextCellinAktKS = akt.Cells[u, j];
                    if (nextCellinAktKS != null && nextCellinAktKS.Value != null && nextCellinAktKS.ToString() != "")
                    {
                        mathesFindCell = RegexReg.ScopeWorkInAktKS.Matches(nextCellinAktKS.Value.ToString());
                    }
                    else continue;
                    if (mathesFindCell.Count > 0)
                    {
                        findCellsColumn = akt.Cells[u, j];
                        break;
                    }
                    if (findCellsColumn != null) break;
                }
            }
            return findCellsColumn;
        }
        //возвращает словарь, где ключ ячейка "Итого по разделу", значение - трудоемкость во разделу
        public static Dictionary<Excel.Range, double> FindForChapter(SmetaForGraf smeta)
        {
            Dictionary<Excel.Range, double> forChapter = new Dictionary<Excel.Range, double>();
            MatchCollection mathes1;
            for (int j = 1; j <= smeta.RangeDoc.Rows.Count; j++)
            {
                Excel.Range nameChapter = smeta.SheetDoc.Cells[j, smeta.KeyNumberPosSmeta.Column];
                if (nameChapter != null && nameChapter.Value2 != null && nameChapter.MergeCells && nameChapter.Value2.ToString() != "")
                {
                    string s = nameChapter.Value.ToString();
                    mathes1 = RegexReg.CellTotalForChapter.Matches(s);
                }
                else continue;
                if (mathes1.Count > 0)
                {
                    Excel.Range trChapter = smeta.SheetDoc.Cells[nameChapter.Row, smeta.KeyTrudozatratSmeta.Column];
                    if (trChapter.Value2 != null && trChapter != null && trChapter.Value.ToString() != "")
                    {
                        double trudChapter = Convert.ToDouble(trChapter.Value2);
                        forChapter.Add(nameChapter, trudChapter);
                    }
                }
            }
            return forChapter;
        }
        //возвращает общую трудоемкость по смете в виде цифры, входной параметр - строковый
        public static double NumeralFromCell(string trudozatrata)
        {
            string summaString = null;
            double trudozatratTotal;
            for (int i = 0; i < trudozatrata.Length; i++)
            {
                if ((trudozatrata[i] >= '0' && trudozatrata[i] <= '9') || trudozatrata[i] == ',' || trudozatrata[i] == '.')
                {
                     summaString += trudozatrata[i];
                }
                }
            if (summaString.Contains("."))
            {
                int index = summaString.IndexOf('.');
                if (index == summaString.Length - 1)
                {
                    summaString = summaString.Remove(summaString.Length - 1, 1);
                    trudozatratTotal = Convert.ToDouble(summaString);
                }
                else
                {
                    summaString = summaString.Replace(".", ",");
                    trudozatratTotal = Convert.ToDouble(summaString);
                }
            }
            else
            {
               trudozatratTotal = Convert.ToDouble(summaString);
            }
            return trudozatratTotal;
        }
        //возвращает лист из ячеек "Раздел такой-то"
        public static List<Excel.Range> FindChapter(SmetaForGraf smeta)
        {
            List<Excel.Range> cellsAllChapter = new List<Excel.Range>();
            MatchCollection mathesChapter;
            for (int j = 1; j <= smeta.RangeDoc.Rows.Count; j++)
            {
                Excel.Range nameChapter = smeta.SheetDoc.Cells[j, smeta.KeyNumberPosSmeta.Column];
                if (nameChapter != null && nameChapter.Value2 != null && nameChapter.MergeCells && nameChapter.Value2.ToString() != "")
                {
                    string stringNameChapter = nameChapter.Value.ToString();
                    mathesChapter = RegexReg.CellOfRazdel.Matches(stringNameChapter);
                }
                else continue;
                if (mathesChapter.Count > 0)
                {
                    cellsAllChapter.Add(nameChapter);
                }
            }
            return cellsAllChapter;
        }
        //возвращает полное имя сметы
        public static void GetNameSmeta(string adresSmeta, out string nameFailSmeta)
        {
            nameFailSmeta = null;
            int numberSlash = 0;
            for (int i = 0; i < adresSmeta.Length; i++)
            {
                if (adresSmeta[i] == '\\') numberSlash = i;
            }
            for (int i = numberSlash + 1; i < adresSmeta.Length - 5; i++)
            {
                nameFailSmeta += adresSmeta[i];
            }
        }
    }
}