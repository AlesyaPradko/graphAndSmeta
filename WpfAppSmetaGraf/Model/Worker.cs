using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public abstract class Worker
    {
        protected string _userAdresSmeta;
        protected string _userAdresKS;
        protected string _userAdresSave;
        protected List<Smeta> _containFolderSmeta;
        protected List<AktKS> _containFolderAktKS;
        protected List<AktKS> _aktKSToOneSmeta;
        protected List<Smeta> _containCopySmeta;
        protected Dictionary<string, List<string>> _aktAllKSforOneSmeta;

      //инициализация адресов папок со сметами, актами КС-2 и куда сохранить ведомость
        public void Initialization(string userSmeta, string userKS, string userWhereSave)
        {
            _userAdresSmeta = userSmeta;
            _userAdresKS = userKS;
            _userAdresSave = userWhereSave;
            _userAdresSave += "\\Копия";
        }
        //возврат листа скопированных смет
        private List<Smeta> MadeCopySmeta()
        {
            List<Smeta> copySmeta = new List<Smeta>();
            for (int u = 0; u < _containFolderSmeta.Count; u++)
            {
                string testuserwheresave = _userAdresSave;
                testuserwheresave += $"{ _containFolderSmeta[u].AddressDoc.Remove(0, _userAdresSmeta.Length + 1)}";//оставляет имя сметы(без пути)      
                Smeta excelBookcopySmet = ParserExcel.CopyExcelSmetaOne(_containFolderSmeta[u], testuserwheresave);
                if (excelBookcopySmet != null)
                { copySmeta.Add(excelBookcopySmet); }
            }
            return copySmeta;
        }
        //если в названии сметы отсутствует №, то смета закрывается
        private void CheckNumber()
        {
            for (int i = 0; i < _containFolderSmeta.Count; i++)
            {
                if (!_containFolderSmeta[i].AddressDoc.Contains("№"))
                {
                    ParserExcel.CloseDoc(_containFolderSmeta[i]);
                    _containFolderSmeta[i].Error += $"В названии сметы {_containFolderSmeta[i].AddressDoc} отсутствует символ № перед номером сметы \n";
                    _containFolderSmeta.RemoveAt(i);
                }
            }
        }
        //первоначальный процесс получения всех листво смет и актов КС-2
        public void ProccessWithDoc(int size,ref string textError)
        {
            _containFolderSmeta = ParserExcel.GetAllSmeta(_userAdresSmeta);
            _containFolderAktKS = ParserExcel.GetAllAktKS(_userAdresKS);
            if (_containFolderSmeta.Count == 0)
            {
                throw new DontHaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку");
            }
            CheckNumber();
            if (_containFolderSmeta.Count == 0)
            {
                CheckIt.Instance.Quit();
                throw new NullValueException("");
            } 
            if (_containFolderAktKS.Count == 0 )
            {
                throw new DontHaveExcelException("В указанной вами папке нет файлов формата .xlsx. Попробуйте выбрать другую папку\n");
            }
            _aktAllKSforOneSmeta = ParserExcel.GetContainAktKSinOneSmeta(_containFolderAktKS, _containFolderSmeta);
            _containCopySmeta = MadeCopySmeta();
            int count = 0;
            List<string> name= new List<string>();
            for (int numSmeta = 0; numSmeta < _containCopySmeta.Count; numSmeta++)
            {
                _aktKSToOneSmeta = GetAllAktToOneSmeta(numSmeta,ref name);
                count += _aktKSToOneSmeta.Count;             
                ProcessSmeta(numSmeta, size,ref textError);
            }
            if(count!= _containFolderAktKS.Count)
            {
                CloseLess(name);
            }         
        }
        //закрытие актов КС-2 в которых отсутствует номер сметы
        private void CloseLess(List<string> name)
        {
            object _misValue = System.Reflection.Missing.Value;
            int count;
            for (int i = 0; i < _containFolderAktKS.Count; i++)
            {
                count = 0;
                for (int j = 0; j < name.Count; j++)
                {
                    if (_containFolderAktKS[i].AddressDoc == name[j]) count++;
                }
                if (count == 0)
                {
                    _containFolderAktKS[i].Error += $"В акте { _containFolderAktKS[i].AddressDoc} отсутсвует номер какой-либо из рассматриваемых смет";
                    _containFolderAktKS[i].DocCur.Close(false, _misValue, _misValue);
                }
            }
        }
        //возвращает лист всех актов КС-2, относящихся к одной смете
        private List<AktKS> GetAllAktToOneSmeta(int numSmeta,ref List<string> name)
        {
            List<AktKS> listAktKStoOneSmeta = new List<AktKS>();
            for (int v = 0; v < _aktAllKSforOneSmeta[_containFolderSmeta[numSmeta].AddressDoc].Count; v++)
            {
                for (int numKS = 0; numKS < _containFolderAktKS.Count; numKS++)
                {
                    if (_containFolderAktKS[numKS].AddressDoc != _aktAllKSforOneSmeta[_containFolderSmeta[numSmeta].AddressDoc][v]) continue;
                    else
                    {
                        listAktKStoOneSmeta.Add(_containFolderAktKS[numKS]);
                        name.Add(_containFolderAktKS[numKS].AddressDoc);
                    }
                }
            }
            return listAktKStoOneSmeta;
        }
        //метод переопределяется в классах-наследниках для работы над сметой в разных режимах
        protected abstract void ProcessSmeta(int num, int size,ref string textError);
        //метод задает формат записи
     
        protected void FormatRecordCopySmeta(Smeta smeta, int size)
        {
            int widthTabl = 0;
            GetNewRange(smeta, ref widthTabl);
            Excel.Range lastCellFormat = smeta.SheetDoc.Cells[smeta.RangeDoc.Rows.Count + smeta.KeyNumberPosSmeta.Row - 1, smeta.RangeDoc.Column + widthTabl - 1];
            Excel.Range firstCellFormat = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, smeta.RangeDoc.Column];
            Excel.Range formarRange = smeta.SheetDoc.get_Range(firstCellFormat, lastCellFormat);
            formarRange.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            formarRange.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            formarRange.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            formarRange.EntireColumn.Font.Size = size;
            formarRange.EntireColumn.Font.FontStyle = "normal";
            formarRange.EntireColumn.AutoFit();
            Excel.Range lastCellwithAnotherWidth = smeta.SheetDoc.Cells[lastCellFormat.Row, smeta.RangeDoc.Column];
            Excel.Range rangewithAnotherWidth = smeta.SheetDoc.get_Range(firstCellFormat, lastCellwithAnotherWidth);
            rangewithAnotherWidth.ColumnWidth = 12;
        }
        //метод изменяет значение количества строк в документе
        private static void GetNewRange(Smeta smeta, ref int widthTabl)
        {
            int testEmptyCells = 0;
            for (int x = smeta.RangeDoc.Column; x < smeta.RangeDoc.Columns.Count + smeta.RangeDoc.Column; x++)
            {
                Excel.Range cellsFirstRowTabl = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, x];
                if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value2 != null && cellsFirstRowTabl.Value2.ToString() != "")
                {
                    widthTabl++;
                    testEmptyCells = 0;
                }
                else
                {
                    testEmptyCells++;
                }
                if (testEmptyCells > 5) break;
            }
        }
    }
}