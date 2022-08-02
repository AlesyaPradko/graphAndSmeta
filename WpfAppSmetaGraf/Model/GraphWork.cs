using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class GraphWork
    {
        private string _userSmeta;
        private Dictionary<int, double> _chelChasForEachWork;
        private Dictionary<int, string> _nameForEachWorkinSmeta;
        private Dictionary<string, List<int>> _dayOnEachWork;
        private List<Dictionary<int, int>> _allChapterInOrder;     
        public int _monthsForWork;
        public string _graphAdress;
        private Dictionary<int, int> _amounWorkInChapter;
        private SmetaForGraf _smeta;
  
        //метод инициализирует адрес сметы и адрес куда требуется сохранить график
        public void InitializationGraph(string userOneSmeta, string userWhereSave)
        {
            _userSmeta = userOneSmeta;
            _graphAdress = userWhereSave;
        }

        //метод для создания сметы и заполнения ее полей
        public void ProccessGraphFirst()
        {
             _smeta= new SmetaForGraf(_userSmeta);
        }
        //метод для получения всех необходимых данных з смеиы для построения графика
        public void ProccessGraph(ref string textError)
        {
            try
            {
                _amounWorkInChapter = new Dictionary<int, int>();
                List<int> deleteChapter = new List<int>();
                _chelChasForEachWork = ChelChasForWorks(_smeta, ref deleteChapter);
                if (_smeta.CellsAllChapter.Count > deleteChapter.Count)
                {
                    GetTakeNewAllChapter(deleteChapter, _smeta);
                }
                _nameForEachWorkinSmeta = NameWorkInPozSmeta(_smeta, ref _amounWorkInChapter);
                int[] startChapterWithWork = _amounWorkInChapter.Keys.ToArray();
                if (_smeta.OnChapterTrudozatrat.Count < _smeta.CellsAllChapter.Count)
                {
                    string er = "Проверьте написание выражений [Итого по разделу]";
                    textError += er;
                    throw new NullValueException(er);
                }
                WorkSorte(_smeta);
                textError += _smeta.Error;
            }
            catch (NullReferenceException ex)
            {
                textError += $"{ex.Message} Проверьте чтобы в {_smeta.AddressDoc} было верно записано устойчивое выражение [№ пп] или [Кол.] или [Т / з осн.раб.Всего]\n";
            }
            catch (InvalidComObjectException ex)
            {
                textError += $"{ex.Message} Проверьте чтобы в {_smeta.AddressDoc} было верно записано устойчивое выражение [№ пп] или [Кол.] или [Т / з осн.раб.Всего]\n";
            }
            catch (COMException ex)
            {
                textError += ex.Message;
            }
        }
       //сортировка сметных работ в порядке их выполнения на стройплощадке 
        private void WorkSorte(SmetaForGraf smeta)
        {
            List<string> ChapterAll = new List<string>();
            List<string> forChapterlAll = new List<string>();
            WorkWithDataBace(ref ChapterAll, ref forChapterlAll);       
            _allChapterInOrder = new List<Dictionary<int, int>>();
            if (forChapterlAll.Count != 0)
            {
                List<List<string>> forChapterAllForRegex = GetListForRegex(forChapterlAll);
                List<Regex> FORChapterAll = new List<Regex>();
                for (int k = 0; k < forChapterAllForRegex.Count; k++)
                {
                    Regex forezd = GetListRegex(forChapterAllForRegex[k]);
                    FORChapterAll.Add(forezd);
                }
                for (int i = 0; i < ChapterAll.Count; i++)
                {
                    RankingAllWorksInOrder(ChapterAll[i], FORChapterAll[i], smeta);
                }
            }
            else
            {
                for (int i = 0; i < ChapterAll.Count; i++)
                {
                    RankingAllWorksInOrder(ChapterAll[i], smeta);
                }
            }
        }
        //сбор информации из базы данных 
    private void WorkWithDataBace(ref List<string> ChapterAll, ref List<string> forChapterlAll)
        {
            string nameSmeta;
            int smetaId = 0;
            using (AppContext db = new AppContext())
            {
                var orderDetails =
                from details in db.Estimates
                where _smeta.AddressDoc.Contains(details.EstimateName)
                select details;
                foreach (var detail in orderDetails)
                {
                    nameSmeta = detail.EstimateName;
                    smetaId = detail.Id;
                    break;
                }
                var orderChap =
                from details in db.Chapters
                where details.EstimateId == smetaId
                select details;
                foreach (var detail in orderChap)
                {
                    ChapterAll.Add(detail.ChapterName);
                    if (detail.WorkName != null)
                    { forChapterlAll.Add(detail.WorkName); }
                    else continue;
                }
            }
        }


        //меняет листы с первыми позициями раздела и листы с ячейками разделов, остаются только те разделы, где есть работа
        private void GetTakeNewAllChapter(List<int> deleteChapter,SmetaForGraf smeta)
        {
            for (int i = smeta.StartChapter.Count - 1; i >= 0; i--)
            {
                int countchap = 0;
                for (int j = deleteChapter.Count - 1; j >= 0; j--)
                {

                    if (i == deleteChapter[j]) countchap++;
                    if (countchap == 1) break;
                }
                if (countchap == 0) 
                {
                    smeta.StartChapter.RemoveAt(i);
                    smeta.CellsAllChapter.RemoveAt(i);
                }
            }
        }
        //возвращает лист с наименованием раздела и работ в этом разделе
        private List<List<string>> GetListForRegex(List<string> forRazdelAll)
        {
            List<List<string>> forRazdelAllForRegex = new List<List<string>>();

            for (int i = 0; i < forRazdelAll.Count; i++)
            {
                string test = forRazdelAll[i];
                List<string> Test = new List<string>();
                string word = null;
                for (int j = 0; j < test.Length; j++)
                {
                    if (test[j] != ',')
                    {
                        word += test[j];
                    }
                    else
                    {
                        Test.Add(word);
                        word = null;
                    }
                }
                Test.Add(word);
                forRazdelAllForRegex.Add(Test);
            }
            return forRazdelAllForRegex;
        }
      
        //возвращает минимальное число рабочих дней
        public int GetMinDays()
        {
            int inputDaysMin = (int)(0.025 * _smeta.TrudozatratTotal / 8);
            if (inputDaysMin == 0) inputDaysMin += 1;
            return inputDaysMin;
        }
        //возвращает максимальное число рабочих дней
        public int GetMaxDays()
        {
            int inputDaysMax = (int)(0.13 * _smeta.TrudozatratTotal / 8);
            return inputDaysMax;
        }
        //возвращает минимальное число рабочих
        public int GetMinPeople()
        {
            int inputWorkersMin = (int)(0.03 * _smeta.TrudozatratTotal / 8);
            if (inputWorkersMin == 0) inputWorkersMin += 1;
            return inputWorkersMin;
        }
        //возвращает максимальное число рабочих
        public int GetMaxPeople()
        {
            int inputWorkersMax = (int)(0.21 * _smeta.TrudozatratTotal / 8);
            return inputWorkersMax;
        }
        //установить число рабочих дней если пользователь выбрал число рабочих
        public void InputDays(ref int daysForWork, ref int amountWorkers)
        {
            double deltaLessThenOne;
            amountWorkers = (int)(_smeta.TrudozatratTotal / (daysForWork * 8));
            deltaLessThenOne = (_smeta.TrudozatratTotal / (daysForWork * 8)) - amountWorkers;
            if (deltaLessThenOne >= 0.5)
            {
                amountWorkers += 1;
                daysForWork += 1;
            }
            else
            {
                daysForWork += 2;
            }
        }
        //установить число рабочих если пользователь выбрал число рабочих дней
        public void InputWorkers(int amountWorkers, ref int daysForWork)
        {
            double deltaLessThenOne;
            daysForWork = (int)(_smeta.TrudozatratTotal / (amountWorkers * 8));
            deltaLessThenOne = (_smeta.TrudozatratTotal / (amountWorkers * 8)) - daysForWork;
            if (deltaLessThenOne > 0.05) daysForWork += 2;
            else daysForWork += 1;
        }
        //возвращает регулярное выражение с видами работ
        private Regex GetListRegex(List<string> forChapter)
        {
            Regex forChapterReg = null;
            switch (forChapter.Count)
            {
                case 1:
                    forChapterReg = new Regex($@"{forChapter[0]}", RegexOptions.IgnoreCase);
                    break;
                case 2:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]})", RegexOptions.IgnoreCase);
                    break;
                case 3:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]})", RegexOptions.IgnoreCase);
                    break;
                case 4:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]})", RegexOptions.IgnoreCase);
                    break;
                case 5:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]})", RegexOptions.IgnoreCase);
                    break;
                case 6:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]})", RegexOptions.IgnoreCase);
                    break;
                case 7:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]})", RegexOptions.IgnoreCase);
                    break;
                case 8:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]})", RegexOptions.IgnoreCase);
                    break;
                case 9:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]})", RegexOptions.IgnoreCase);
                    break;
                case 10:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]})", RegexOptions.IgnoreCase);
                    break;
                case 11:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]})", RegexOptions.IgnoreCase);
                    break;
                case 12:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]})", RegexOptions.IgnoreCase);
                    break;
                case 13:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]})", RegexOptions.IgnoreCase);
                    break;
                case 14:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]})", RegexOptions.IgnoreCase);
                    break;
                case 15:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]})", RegexOptions.IgnoreCase);
                    break;
                case 16:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]})", RegexOptions.IgnoreCase);
                    break;
                case 17:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]})", RegexOptions.IgnoreCase);
                    break;
                case 18:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]})", RegexOptions.IgnoreCase);
                    break;
                case 19:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]})", RegexOptions.IgnoreCase);
                    break;
                case 20:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]})", RegexOptions.IgnoreCase);
                    break;
                case 21:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]})", RegexOptions.IgnoreCase);
                    break;
                case 22:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]})", RegexOptions.IgnoreCase);
                    break;
                case 23:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]}|{forChapter[22]})", RegexOptions.IgnoreCase);
                    break;
                case 24:
                    forChapterReg = new Regex($@"({forChapter[0]}|{forChapter[1]}|{forChapter[2]}|{forChapter[3]}|{forChapter[4]}|{forChapter[5]}|{forChapter[6]}|{forChapter[7]}|{forChapter[8]}|{forChapter[9]}|{forChapter[10]}|{forChapter[11]}|{forChapter[12]}|{forChapter[13]}|{forChapter[14]}|{forChapter[15]}|{forChapter[16]}|{forChapter[17]}|{forChapter[18]}|{forChapter[19]}|{forChapter[20]}|{forChapter[21]}|{forChapter[22]}|{forChapter[23]})", RegexOptions.IgnoreCase);
                    break;
            }
            return forChapterReg;
        }


        //возвращает  словарь, где ключ - номер по смете, значение - трудозатраты на данную работу
        private Dictionary<int, double> ChelChasForWorks(SmetaForGraf smeta, ref List<int> deleteChapter)
        {
            Dictionary<int, double> chelChasforEachWork = new Dictionary<int, double>();
            double trudozatratOfWork;
            int numPosSmeta;
            for (int j = smeta.KeyNumberPosSmeta.Row + 4; j <= smeta.RangeDoc.Rows.Count; j++)
            {
                Excel.Range cellsNumberPosColumnTabl = smeta.SheetDoc.Cells[j, smeta.KeyNumberPosSmeta.Column];
                Excel.Range cellsColumnTrudozatrat = smeta.SheetDoc.Cells[j, smeta.KeyTrudozatratSmeta.Column];
                if (cellsNumberPosColumnTabl != null && cellsNumberPosColumnTabl.Value2 != null && !cellsNumberPosColumnTabl.MergeCells && cellsNumberPosColumnTabl.Value2.ToString() != "" && cellsColumnTrudozatrat != null && cellsColumnTrudozatrat.Value2 != null && !cellsColumnTrudozatrat.MergeCells && cellsColumnTrudozatrat.Value2.ToString() != "")
                {
                    try
                    {
                        int numCellsForNumPosSmeta = cellsNumberPosColumnTabl.Row;
                        numPosSmeta = Convert.ToInt32(cellsNumberPosColumnTabl.Value2);
                        trudozatratOfWork = Convert.ToDouble(cellsColumnTrudozatrat.Value2);
                        chelChasforEachWork.Add(numPosSmeta, trudozatratOfWork);
                        for (int i = 0; i < smeta.StartChapter.Count; i++)
                        {
                            if (smeta.StartChapter[i] == numPosSmeta)
                            {
                                deleteChapter.Add(i);
                            }

                        }
                    }
                    catch (NullReferenceException ex)
                    {
                        smeta.Error += $"{ex.Message} Проверьте чтобы в {smeta.AddressDoc} было верно записано устойчивое выражение [Наименование]\n";
                    }
                    catch (ArgumentException ex)
                    {
                        smeta.Error += $"{ex.Message} Проверьте чтобы в {smeta.AddressDoc} не повторялись значения позиций по смете в строке {cellsNumberPosColumnTabl.Row}\n";
                    }
                    catch (FormatException ex)
                    {
                        smeta.Error += $"{ex.Message} Вы ввели неверный формат для {smeta.AddressDoc} в строке {cellsNumberPosColumnTabl.Row} в столбце {cellsNumberPosColumnTabl.Column}(не должно быть [., букв], только целые числа,или в столбце {cellsColumnTrudozatrat.Column} только числа дробные, не должно быть [.букв]  )\n";
                    }
                }
            }
            return chelChasforEachWork;
        }
        //возвращает  словарь, где ключ - номер по смете, значение - строковое наименование данных работ
        private Dictionary<int, string> NameWorkInPozSmeta(SmetaForGraf smeta, ref Dictionary<int, int> _amounWorkInChapter)
        {
            int[] keyTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            Dictionary<int, string> nameForEachWorkInSmeta = new Dictionary<int, string>();
            int numPosSmeta;
            int count = 0;
            string nameWorkInPosSmeta;
            for (int j = smeta.KeyNumberPosSmeta.Row + 4; j <= smeta.RangeDoc.Rows.Count; j++)
            {
                 Excel.Range cellsNumberPosColumnTabl = smeta.SheetDoc.Cells[j, smeta.KeyNumberPosSmeta.Column];
                 Excel.Range cellsNameWorkColumnTabl = smeta.SheetDoc.Cells[j, smeta.KeyCellNameWork.Column];
                 if (cellsNumberPosColumnTabl != null && cellsNumberPosColumnTabl.Value2 != null && !cellsNumberPosColumnTabl.MergeCells && cellsNumberPosColumnTabl.Value2.ToString() != "" && cellsNameWorkColumnTabl != null && cellsNameWorkColumnTabl.Value2 != null && !cellsNameWorkColumnTabl.MergeCells && cellsNameWorkColumnTabl.Value2.ToString() != "")
                 {
                     try
                     {
                        for (int i = 0; i < keyTrudozatratEachWork.Length; i++)
                        {
                            numPosSmeta = Convert.ToInt32(cellsNumberPosColumnTabl.Value2);
                            if (numPosSmeta == keyTrudozatratEachWork[i])
                            {
                                nameWorkInPosSmeta = cellsNameWorkColumnTabl.Value.ToString();
                                nameForEachWorkInSmeta.Add(numPosSmeta, nameWorkInPosSmeta);
                                for (int k = 0; k < smeta.CellsAllChapter.Count; k++)
                                {
                                    if (numPosSmeta == smeta.StartChapter[k])
                                    {
                                        if (i == 0) { count = 0; }
                                        else
                                        {
                                           _amounWorkInChapter.Add(smeta.StartChapter[k - 1], count);
                                           count = 0;
                                        }
                                    }
                                }
                                count++;
                            }
                        }
                     }
                     catch (ArgumentException ex)
                     {
                            smeta.Error += $"{ex.Message} Проверьте чтобы в {smeta.AddressDoc} не повторялись значения позиций по смете в строке {cellsNumberPosColumnTabl.Row}\n";
                     }
                     catch (FormatException ex)
                     {
                        smeta.Error += $"{ex.Message} Вы ввели неверный формат для {smeta.AddressDoc} в строке {cellsNumberPosColumnTabl.Row} в столбце {cellsNumberPosColumnTabl.Column}(не должно быть [., букв], только целые числа.\n";
                     }
                 }
            }    
            if (count != 0) { _amounWorkInChapter.Add(smeta.StartChapter[smeta.StartChapter.Count - 1], count); }
            return nameForEachWorkInSmeta;
        }

        //меняет по ссылке лист, состоящий из словарей,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех разделов
        private void RankingAllWorksInOrder(string regulNameOfRazdel, Regex regulNameWorkOfRazdel, SmetaForGraf smeta)
        {
            int[] keyNumTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            string[] valueNameofEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            Dictionary<int, int> inChapterNumPosAndNumWorkInArr;
            for (int i = 0; i < smeta.CellsAllChapter.Count; i++)
            {
                string stringPoRazdelyforPoisk = smeta.CellsAllChapter[i].Value.ToString();
                if (stringPoRazdelyforPoisk.Contains(regulNameOfRazdel))
                {
                    if (i < smeta.CellsAllChapter.Count - 1)
                    {
                        inChapterNumPosAndNumWorkInArr = InOrderChapter(regulNameWorkOfRazdel, valueNameofEachWork, keyNumTrudozatratEachWork, smeta.StartChapter[i], smeta.StartChapter[i + 1]);
                        GetOrderChapter(inChapterNumPosAndNumWorkInArr);
                    }
                    else
                    {
                        inChapterNumPosAndNumWorkInArr = InOrderChapter(regulNameWorkOfRazdel, valueNameofEachWork, keyNumTrudozatratEachWork, smeta.StartChapter[i]);
                        GetOrderChapter(inChapterNumPosAndNumWorkInArr);
                    }
                }
            }
        }
        //меняет по ссылке лист, состоящий из словарей,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех разделов
        private void RankingAllWorksInOrder(string regulNameOfRazdel,SmetaForGraf smeta)
        {
            int[] keyNumTrudozatratEachWork = _chelChasForEachWork.Keys.ToArray();
            string[] valueNameofEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            Dictionary<int, int> inChapterNumPosAndNumWorkInArr;
            for (int i = 0; i < smeta.CellsAllChapter.Count; i++)
            {
                string stringPoRazdelyforPoisk = smeta.CellsAllChapter[i].Value.ToString();
                if (stringPoRazdelyforPoisk.Contains(regulNameOfRazdel))
                {
                    if (i < smeta.CellsAllChapter.Count - 1)
                    {
                        inChapterNumPosAndNumWorkInArr = InOrderChapter(valueNameofEachWork, keyNumTrudozatratEachWork, smeta.StartChapter[i], smeta.StartChapter[i + 1]);
                        GetOrderChapter(inChapterNumPosAndNumWorkInArr);
                    }
                    else
                    {
                        inChapterNumPosAndNumWorkInArr = InOrderChapter(valueNameofEachWork, keyNumTrudozatratEachWork, smeta.StartChapter[i]);
                        GetOrderChapter(inChapterNumPosAndNumWorkInArr);
                    }
                }
            }
        }
        private void GetOrderChapter(Dictionary<int, int> inChapterNumPosAndNumWorkInArr)
        {
            if (inChapterNumPosAndNumWorkInArr.Count > 0)
            {
                _allChapterInOrder.Add(inChapterNumPosAndNumWorkInArr);
            }
        }
        //возвращает словарь,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех разделов
        private Dictionary<int, int> InOrderChapter(Regex regulNameWorkOfChapter, string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork, int startChapt)
        {
            MatchCollection mathesNameWork;
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                mathesNameWork = regulNameWorkOfChapter.Matches(valueNameOfEachWork[i]);
                if (mathesNameWork.Count > 0)
                {
                    if (keyNumberTrudozatratEachWork[i] >= startChapt)
                    {
                        inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                    }
                }
            }
            return inChapterNumberPosAndNumWorkInArr;
        }
        //возвращает словарь,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех
        private Dictionary<int, int> InOrderChapter(string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork, int startChapt)
        {
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                if (keyNumberTrudozatratEachWork[i] >= startChapt)
                {
                    inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                }
            }
            return inChapterNumberPosAndNumWorkInArr;
        }

        //возвращает словарь,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех
        private Dictionary<int, int> InOrderChapter(Regex regulNameWorkOfChapter, string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork, int startChapt, int lastChapt)
        {
            MatchCollection mathesNameWork;
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                mathesNameWork = regulNameWorkOfChapter.Matches(valueNameOfEachWork[i]);
                if (mathesNameWork.Count > 0)
                {
                    if (keyNumberTrudozatratEachWork[i] >= startChapt && keyNumberTrudozatratEachWork[i] < lastChapt)
                    {
                        inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                    }
                }
            }
            return inChapterNumberPosAndNumWorkInArr;
        }
        //возвращает словарь,где ключ - номер по смете, значение - номер строки в массиве наименования данных работ для всех
        private Dictionary<int, int> InOrderChapter(string[] valueNameOfEachWork, int[] keyNumberTrudozatratEachWork, int startChapt, int lastChapt)
        {
            Dictionary<int, int> inChapterNumberPosAndNumWorkInArr = new Dictionary<int, int>();
            for (int i = 0; i < valueNameOfEachWork.Length; i++)
            {
                if (keyNumberTrudozatratEachWork[i] >= startChapt && keyNumberTrudozatratEachWork[i] < lastChapt)
                {
                    inChapterNumberPosAndNumWorkInArr.Add(keyNumberTrudozatratEachWork[i], i);
                }

            }
            return inChapterNumberPosAndNumWorkInArr;
        }

        //закрашивает график в соответствие с данными
        public void RecordGraph(DateTime dataStart, int daysForWork, int numberofWorkers, int color, ref string _textError)
        {
            try
            {
                string nameFailSmeta;
                ParserExcel.GetNameSmeta(_smeta.AddressDoc, out nameFailSmeta);
                _graphAdress += $"\\График производства работ - {nameFailSmeta}";
                Excel.Workbook workBookGraph = CheckIt.Instance.Workbooks.Add();
                Excel.Worksheet workSheetGraph = (Excel.Worksheet)workBookGraph.Worksheets.get_Item(1);
                Excel.Range FirstCellGraph = workSheetGraph.Range["B4"];
                Excel.Range lastMonth = null;
                RecordNotes(workSheetGraph, daysForWork, dataStart, ref lastMonth);
                int amountOfWorkInChapter = 0;
                int[] numChapterTablExcelGraph = new int[_allChapterInOrder.Count];
                int[] amountOfWorkersInChapter = new int[_allChapterInOrder.Count]; ;
                RecordAllString(numberofWorkers, workSheetGraph, ref amountOfWorkInChapter, ref numChapterTablExcelGraph, ref amountOfWorkersInChapter);
                Excel.Range LastCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + amountOfWorkInChapter + 1, lastMonth.Column];
                Excel.Range rangeGraph = workSheetGraph.get_Range(FirstCellGraph, LastCellGraph);
                rangeGraph.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                rangeGraph.EntireColumn.Font.Size = 10;
                rangeGraph.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
                rangeGraph.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
                rangeGraph.EntireColumn.AutoFit();
                Excel.Range cellforDaysSimilarSize = workSheetGraph.get_Range("G5", LastCellGraph);
                cellforDaysSimilarSize.ColumnWidth = 4;
                FillColor(LastCellGraph, numChapterTablExcelGraph, workSheetGraph, color);
                FirstCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + 2, FirstCellGraph.Column + 1];
                LastCellGraph = workSheetGraph.Cells[FirstCellGraph.Row + amountOfWorkInChapter + 1, FirstCellGraph.Column + 1];
                Excel.Range rangeCellsGrafik = workSheetGraph.get_Range(FirstCellGraph, LastCellGraph);
                rangeCellsGrafik.EntireColumn.HorizontalAlignment = Excel.Constants.xlLeft;
                workBookGraph.SaveAs(_graphAdress);
                object misValue = System.Reflection.Missing.Value;
                _smeta.DocCur.Close(false, misValue, misValue);
                workBookGraph.Close(true, misValue, misValue);
            }
            catch (COMException exc)
            {
                _textError += $"{exc.Message} Закройте файл графика и повторите снова/Файл не будет пересохранен";
                return;
            }
            catch (NullReferenceException exc)
            {
                _textError += $"{exc.Message} Проверьте правильность написания Сметной трудоемкости";
                return;
            }
            catch (InvalidComObjectException exc)
            {
                _textError += $"{exc.Message} Закройте файл графика и повторите снова";
                return;
            }
        }
        //получение из базы данных всех дат рабочих дней на производство работ
        private DateTime[] FindDayMonths(DateTime dataStartWork, int daysForWork)
        {
            DateTime[] workDays = new DateTime[daysForWork];
            using (AppContext db = new AppContext())
            {
                var orderDetails =
                from details in db.WorkingDays
                where details.Date >= dataStartWork
                select details;
                int i = 0;
                foreach (var detail in orderDetails)
                {
                    workDays[i] = detail.Date;
                    i++;
                    if (i == daysForWork) break;
                }
            }
            return workDays;
        }
        //возвращает словарь,где ключ - название месяца, значение - лист рабочих дней месяца 
        public Dictionary<string, List<int>> GetDaysForWork(DateTime[] workDays)
        {
            Dictionary<string, List<int>> dayOnEachWork = new Dictionary<string, List<int>>();
            List<int> dayMonth = new List<int>();
            string text = ParserExcel.MonthLetterInt(workDays[0].Month) + '.' + workDays[0].Year.ToString();
            for (int i = 0; i < workDays.Length; i++)
            {
                if (i == 0) dayMonth.Add(workDays[i].Day);
                else
                {
                    if (workDays[i].Day > workDays[i - 1].Day)
                    {
                        dayMonth.Add(workDays[i].Day);
                    }

                    if (workDays[i].Day < workDays[i - 1].Day)
                    {
                        dayOnEachWork.Add(text, dayMonth);
                        text = ParserExcel.MonthLetterInt(workDays[i].Month) + '.' + workDays[i].Year.ToString();
                        dayMonth = new List<int>();
                        dayMonth.Add(workDays[i].Day);
                    }
                }
            }
            if (dayMonth.Count > 0)
            {
                dayOnEachWork.Add(text, dayMonth);
            }
            return dayOnEachWork;
        }
        //заполнение шапки и данных таблицы графика
        private void RecordNotes(Excel.Worksheet workSheetGraph, int daysForWork, DateTime dataStartWork, ref Excel.Range lastMonth)
        {
            Excel.Range GraphNext = workSheetGraph.get_Range("B4", "B5");
            GraphNext.Merge();
            GraphNext.Value = "№";
            GraphNext = workSheetGraph.get_Range("C4", "C5");
            GraphNext.Merge();
            GraphNext.Value = "Наименование работ";
            GraphNext = workSheetGraph.get_Range("D4", "D5");
            GraphNext.Merge();
            GraphNext.Value = "Всего чел/час";
            GraphNext = workSheetGraph.get_Range("E4", "E5");
            GraphNext.Merge();
            GraphNext.Value = "Кол. чел.  бр";
            GraphNext = workSheetGraph.get_Range("F4", "F5");
            GraphNext.Merge();
            GraphNext.Value = "Кол-во рабоч. дней";
            Excel.Range firstMonth;
            double delta = (daysForWork / 21.0) - (int)(daysForWork / 21);
            if (delta < 0.04) _monthsForWork = daysForWork / 21;
            else _monthsForWork = 1 + daysForWork / 21;
            DateTime[] workDays = FindDayMonths(dataStartWork, daysForWork);
            _dayOnEachWork = GetDaysForWork(workDays);
            List<int>[] valueAllWorkDaysForMonth = _dayOnEachWork.Values.ToArray();
            string[] keyNameDataWork = _dayOnEachWork.Keys.ToArray();
            for (int i = 0; i < valueAllWorkDaysForMonth.Length; i++)
            {
                firstMonth = workSheetGraph.Cells[GraphNext.Row, GraphNext.Column + 1];
                lastMonth = workSheetGraph.Cells[GraphNext.Row, GraphNext.Column + valueAllWorkDaysForMonth[i].Count];
                for (int j = 0; j < valueAllWorkDaysForMonth[i].Count; j++)
                {
                    workSheetGraph.Cells[firstMonth.Row + 1, firstMonth.Column + j] = valueAllWorkDaysForMonth[i][j];
                }
                GraphNext = workSheetGraph.get_Range(firstMonth, lastMonth);
                GraphNext.Merge();
                GraphNext.Value = keyNameDataWork[i];
                GraphNext = lastMonth;
            }

        }
        //закрашивание цветом границ графика
        private void FillColor(Excel.Range LastCellGraph, int[] numChapterTablExcelGraph, Excel.Worksheet workSheetGraph, int color)
        {
            int amountOfDaysOnAllChapter = 0, summaAmountofDaysEachWork = 0, summaAmountofWorkerEachWork, indexofChapter = 0;
            int amountOfDaysOnEachChapter, amountofWorkerOnEachChapter;
            double deltaWorker = 0;
            double allWorker, sumAllWork = 0;
            Excel.Range rangeForColour = workSheetGraph.get_Range("E6", LastCellGraph);
            for (int j = rangeForColour.Row; j < rangeForColour.Rows.Count + rangeForColour.Row; j++)
            {
                if (indexofChapter < numChapterTablExcelGraph.Length)
                {
                    if (j == numChapterTablExcelGraph[indexofChapter])
                    {
                        indexofChapter++;
                        Excel.Range amountofDaysEachRazdelTabl = workSheetGraph.Cells[numChapterTablExcelGraph[indexofChapter - 1], 6];
                        amountOfDaysOnEachChapter = (int)(amountofDaysEachRazdelTabl.Value2);
                        amountOfDaysOnAllChapter += amountOfDaysOnEachChapter;
                    }
                }
                if (indexofChapter > 0)
                {
                    if (j >= numChapterTablExcelGraph[indexofChapter - 1] + 1)
                    {
                        Excel.Range amountofWorkerEachRazdelTabl = workSheetGraph.Cells[numChapterTablExcelGraph[indexofChapter - 1], 5];
                        amountofWorkerOnEachChapter = (int)(amountofWorkerEachRazdelTabl.Value2);
                        Excel.Range trudWork = workSheetGraph.Cells[j, 4];
                        Excel.Range numberofWorkerEachWorkTabl = workSheetGraph.Cells[j, 5];
                        Excel.Range numberofDaysEachWorkTabl = workSheetGraph.Cells[j, 6];
                        allWorker = (int)(numberofWorkerEachWorkTabl.Value2) + deltaWorker;
                        sumAllWork += allWorker;
                        summaAmountofWorkerEachWork = (int)sumAllWork;
                        deltaWorker = trudWork.Value2 / (8 * (int)numberofDaysEachWorkTabl.Value2) - (int)(numberofWorkerEachWorkTabl.Value2);
                        if (summaAmountofWorkerEachWork < amountofWorkerOnEachChapter)
                        {
                            Excel.Range firstFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            if (lastFillColour.Column > LastCellGraph.Column) lastFillColour = LastCellGraph;
                            Excel.Range rangeFillColour = workSheetGraph.get_Range(firstFillColour, lastFillColour);
                            rangeFillColour.Interior.ColorIndex = color;
                        }
                        else
                        {
                            Excel.Range firstFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork];
                            Excel.Range lastFillColour;
                            double chekWorker = 1.0 * summaAmountofWorkerEachWork / amountofWorkerOnEachChapter;
                            if (chekWorker >= 1.5)
                                lastFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2)];
                            else
                                lastFillColour = workSheetGraph.Cells[j, 7 + summaAmountofDaysEachWork + (int)(numberofDaysEachWorkTabl.Value2) - 1];
                            if (lastFillColour.Column > LastCellGraph.Column) lastFillColour = LastCellGraph;
                            Excel.Range rangeFillColour = workSheetGraph.get_Range(firstFillColour, lastFillColour);
                            rangeFillColour.Interior.ColorIndex = color;
                            summaAmountofDaysEachWork += (int)(numberofDaysEachWorkTabl.Value2);
                            if (summaAmountofDaysEachWork > amountOfDaysOnAllChapter) summaAmountofDaysEachWork -= 1; //бригада переходит на следующие работы в тот же день                          
                        }
                        if (sumAllWork > amountofWorkerOnEachChapter)
                            sumAllWork -= amountofWorkerOnEachChapter;
                    }
                }
            }
        }
        //заполнение данных по работам таблицы графика
        public void RecordAllString(int amountOfWorkers, Excel.Worksheet workSheetGrafik, ref int amountOfWorkInChapter, ref int[] numChapterTablExcelGrafik, ref int[] amountOfWorkersInChapter)
        {
            Excel.Range firstCellAfterContent = workSheetGrafik.Range["B6"];
            int indexAmountWorkInChapter = 0, AmountofWorkerinEachWork = 0, numPosGrafik = 0, indexChapt = 0;
            double reservPartOfDayAfterWork = 0;
            string[] valueNameOfEachWork = _nameForEachWorkinSmeta.Values.ToArray();
            double[] valueTrudozatratEachWork = _chelChasForEachWork.Values.ToArray();
            for (int i = 0; i < _allChapterInOrder.Count; i++)
            {
                int indexAmountOfRowEachWorkinChapter;
                int[] keyNumPosSmetaChapterInOrder = _allChapterInOrder[i].Keys.ToArray();
                int[] valueNumPosWorkChapterInOrder = _allChapterInOrder[i].Values.ToArray();
                int daysOfEachWork;
                for (int r = 0; r < _smeta.CellsAllChapter.Count; r++)
                {
                    indexAmountOfRowEachWorkinChapter = 0;

                    if (keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] == _smeta.StartChapter[r] && _amounWorkInChapter[_smeta.StartChapter[r]] == keyNumPosSmetaChapterInOrder.Length)
                    {
                        indexAmountOfRowEachWorkinChapter = 0;
                        RecordPartString(workSheetGrafik, firstCellAfterContent, r, amountOfWorkInChapter, ref numChapterTablExcelGrafik, ref numPosGrafik, ref indexChapt);
                        RecordTwoPartString(workSheetGrafik, amountOfWorkers, firstCellAfterContent, r, amountOfWorkInChapter);
                    }
                    else if ((keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] == _smeta.StartChapter[r] && _amounWorkInChapter[_smeta.StartChapter[r]] > keyNumPosSmetaChapterInOrder.Length) || ((r == _smeta.CellsAllChapter.Count - 1 || keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] < _smeta.StartChapter[r + 1]) && keyNumPosSmetaChapterInOrder[indexAmountOfRowEachWorkinChapter] > _smeta.StartChapter[r] && _amounWorkInChapter[_smeta.StartChapter[r]] > keyNumPosSmetaChapterInOrder.Length))
                    {
                        indexAmountOfRowEachWorkinChapter = 0;
                        RecordPartString(workSheetGrafik, firstCellAfterContent, r, amountOfWorkInChapter, ref numChapterTablExcelGrafik, ref numPosGrafik, ref indexChapt);
                        double trudPartOfChapter = 0;
                        for (int q = 0; q < keyNumPosSmetaChapterInOrder.Length; q++)
                        {
                            trudPartOfChapter += _chelChasForEachWork[keyNumPosSmetaChapterInOrder[q]];
                        }
                        RecordTwoPartString(workSheetGrafik, amountOfWorkers, firstCellAfterContent, r, amountOfWorkInChapter, trudPartOfChapter);
                    }
                    else continue;
                    do
                    {
                        if (r < _smeta.CellsAllChapter.Count - 1 && keyNumPosSmetaChapterInOrder[indexAmountWorkInChapter] >= _smeta.StartChapter[r + 1])
                        {
                            break;
                        }
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column] = ++numPosGrafik;
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 1] = valueNameOfEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]];
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 2] = valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]];
                        int amountOfWorkersInOneTime = 0;
                        do
                        {
                            amountOfWorkersInOneTime++;
                            if (valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] > 8 * amountOfWorkers)
                            {
                                AmountofWorkerinEachWork = amountOfWorkers;
                                break;
                            }
                            if (valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] <= 8 * amountOfWorkersInOneTime)
                            {
                                AmountofWorkerinEachWork = amountOfWorkersInOneTime;
                                break;
                            }
                        } while (amountOfWorkersInOneTime <= amountOfWorkers);
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 3] = AmountofWorkerinEachWork;
                        daysOfEachWork = (int)(valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] / (AmountofWorkerinEachWork * 8));
                        reservPartOfDayAfterWork += valueTrudozatratEachWork[valueNumPosWorkChapterInOrder[indexAmountWorkInChapter]] / (AmountofWorkerinEachWork * 8) - daysOfEachWork;
                        if (reservPartOfDayAfterWork >= 1)
                        {
                            daysOfEachWork += 1;
                            reservPartOfDayAfterWork -= 1;
                        }
                        if (daysOfEachWork == 0)
                        {
                            daysOfEachWork += 1;
                        }
                        workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter + indexAmountOfRowEachWorkinChapter + 1, firstCellAfterContent.Column + 4] = daysOfEachWork;
                        indexAmountWorkInChapter++;
                        indexAmountOfRowEachWorkinChapter++;
                        if (indexAmountWorkInChapter == valueNumPosWorkChapterInOrder.Length)
                        {
                            indexAmountWorkInChapter = 0;
                            break;
                        }
                    } while (indexAmountWorkInChapter > 0);
                    amountOfWorkInChapter += indexAmountOfRowEachWorkinChapter + 1;
                }
            }
        }
        //запись раздела
        private void RecordPartString(Excel.Worksheet workSheetGrafik, Excel.Range firstCellAfterContent, int r, int amountOfWorkInChapter, ref int[] numChapterTablExcelGrafik, ref int numPosGrafik, ref int indexChapt)
        {

            string nameOfRazdel = _smeta.CellsAllChapter[r].Value.ToString();
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column] = ++numPosGrafik;
            numChapterTablExcelGrafik[indexChapt++] = firstCellAfterContent.Row + amountOfWorkInChapter;
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 1] = nameOfRazdel;
        }
        //запись работы
        private void RecordTwoPartString(Excel.Worksheet workSheetGrafik, int amountOfWorkers, Excel.Range firstCellAfterContent, int r, int amountOfWorkInChapter)
        {
            double[] trudozatratForChapter = _smeta.OnChapterTrudozatrat.Values.ToArray();
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 2] = trudozatratForChapter[r];
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 3] = amountOfWorkers;
            int daysOfEachWork = (int)(trudozatratForChapter[r] / (amountOfWorkers * 8));
            if (trudozatratForChapter[r] / (amountOfWorkers * 8) - daysOfEachWork > 0.05)
            {
                daysOfEachWork += 1;
            }
            if (daysOfEachWork == 0)
            {
                daysOfEachWork += 1;
            }
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 4] = daysOfEachWork;
        }
        //запись работы
        private void RecordTwoPartString(Excel.Worksheet workSheetGrafik, int amountOfWorkers, Excel.Range firstCellAfterContent, int r, int amountOfWorkInChapter, double trudPartOfChapter)
        {
            double[] trudozatratForChapter = _smeta.OnChapterTrudozatrat.Values.ToArray();
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 2] = trudPartOfChapter;
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 3] = amountOfWorkers;
            int daysOfEachWork = (int)(trudozatratForChapter[r] / (amountOfWorkers * 8));
            if (trudPartOfChapter / (amountOfWorkers * 8) - daysOfEachWork > 0.05)
            {
                daysOfEachWork += 1;
            }
            if (daysOfEachWork == 0)
            {
                daysOfEachWork += 1;
            }
            workSheetGrafik.Cells[firstCellAfterContent.Row + amountOfWorkInChapter, firstCellAfterContent.Column + 4] = daysOfEachWork;
        }

    }
}
