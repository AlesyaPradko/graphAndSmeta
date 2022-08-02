using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace WpfAppSmetaGraf.Model
{
    public class AktKS : DocumentExcel
    {

        private readonly Excel.Range _keyNumberPosKS;
        private readonly Excel.Range _keyScopeWorkinAktKS;
        private readonly string _numAktKS;
        private readonly string _dataAktKS;
        private readonly Dictionary<int, double> _totalScopeWorkAktKSone;
        int _monthAktKS;
        int _yearAktKS;
        int _dayAktKS;
        public Excel.Range KeyNumberPosKS { get { return _keyNumberPosKS; } }
        public string NumAktKS { get { return _numAktKS; } }
        public string DatAktKS { get { return _dataAktKS; } }
        public int MonthAktKS { get { return _monthAktKS; } }
        public int YearAktKS { get { return _yearAktKS; } }
        public int DayAktKS { get { return _dayAktKS; } }

        public Excel.Range KeyScopeWorkinAktKS { get { return _keyScopeWorkinAktKS; } }
        public Dictionary<int, double> TotalScopeWorkAktKSone { get { return _totalScopeWorkAktKSone; } }
        public AktKS(string name) : base(name)
        {
            _keyNumberPosKS = FindText("по смете", this, RangeDoc); ;
            _keyScopeWorkinAktKS = FindText("количество", this, RangeDoc);
            _totalScopeWorkAktKSone = ParserExcel.GetScopeWorkAktKSone(this);
            _numAktKS = FindDataorTime("Номер документа");
            _dataAktKS = FindDataorTime("Дата составления");
            _dayAktKS = ReturnNumberofDate(RegexReg.RegexDay.Matches(DatAktKS), 0);
            _monthAktKS = ReturnNumberofDate(RegexReg.RegexMonth.Matches(DatAktKS), 1);
            _yearAktKS = ReturnNumberofDate(RegexReg.RegexYear.Matches(DatAktKS), 2);
        }
        //находит дату акста КС-2 и его номер
        private string FindDataorTime(string test)
        {
            string result;
            Excel.Range num = RangeDoc.Find(test);
            if (num != null)
            {
                Excel.Range numAkt = SheetDoc.Cells[num.Row + 2, num.Column];
                result = numAkt.Value.ToString();
            }
            else
            {
                string er = $" Проверьте чтобы в {AddressDoc} было верно записано устойчивое выражение [Номер документа] или [Дата составления]\n";
                Error += er;
                ParserExcel.CloseDoc(this);
                throw new NullValueException(er);
            }
            return result;
        }
        //возврат целочисленного значения - день, месяц, год
        private int ReturnNumberofDate(MatchCollection math, int num)
        {
            string findPart = null;
            if (math.Count > 0)
            {
                foreach (Match part in math)
                {
                    findPart = part.Value;
                    findPart = FindDate(num, findPart);
                }
            }
            if (findPart == null)
            {
                string er = $"В акте {AddressDoc} не прописана дата, устраните замечание";
                Error += er;
                throw new NullValueException(er);
            }
            return Convert.ToInt32(findPart);
        }
        //находит строковое выражение дня,месяца, года
        private static string FindDate(int num, string findPart)
        {
            switch (num)
            {
                case 0:
                    findPart = findPart.Remove(findPart.Length - 1, 1);
                    break;
                case 1:
                    findPart = findPart.Remove(findPart.Length - 1, 1);
                    findPart = findPart.Remove(0, 1);
                    break;
                case 2:
                    findPart = findPart.Remove(0, 1);
                    break;
            }

            return findPart;
        }
    }
}
