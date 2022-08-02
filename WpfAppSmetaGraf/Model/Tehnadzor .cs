using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class Tehnadzor:Worker
    {
        private List<AktKS> _aktKSinOrderSort;
        private List<Dictionary<int, double>> _forRecordWorkColumnInSmeta;
        private List<string> _nameAktKSRecordColumn;

        //выполняет работу по сбору информации и построению ведомости
        protected override void ProcessSmeta(int num, int size, ref string textError)
        {
                int nextInsertColumn = _containCopySmeta[num].KeyConstructWorkSmeta.Column + 1;
                int lastRowCellsAfterDelete = 0;
                ParserExcel.DeleteColumnandRow(_containCopySmeta[num], ref lastRowCellsAfterDelete);
                Excel.Range newLastCell = _containCopySmeta[num].SheetDoc.Cells[lastRowCellsAfterDelete, _containCopySmeta[num].RangeDoc.Columns.Count];
                _containCopySmeta[num].RangeDoc = _containCopySmeta[num].SheetDoc.get_Range(_containCopySmeta[num].KeyNumberPosSmeta, newLastCell);//уменьшение области обработки 
                string error = null;
                if (_aktKSToOneSmeta.Count != 0)
                {
                    _aktKSinOrderSort = new List<AktKS>();
                    SortAktKSforTehnadzor();
                    WorkWithListKSTeh(ref error);
                    _containCopySmeta[num].Error += error;
                    for (int i = 0; i < _forRecordWorkColumnInSmeta.Count; i++)
                    {
                        RecordPartFileTehnadzor(i, _containCopySmeta[num], ref nextInsertColumn);
                    }
                    FormatRecordCopySmeta(_containCopySmeta[num], size);
                    ParserExcel.CloseAktKS(_aktKSToOneSmeta);
                }
                if (_aktKSinOrderSort.Count != 0)
                {
                    RecordFormulaTehnadzor(_containCopySmeta[num], nextInsertColumn);
                }
                textError += _containCopySmeta[num].Error;
                ParserExcel.CloseDoc(_containCopySmeta[num]);

        }

        //метод возврашает строку - наименование столбца выполненных объемов работ по КС-2 за определенный период и заполняет словарь
        //где ключ -номер позиции по смете из Актов КС, значение выполнение по смете
        private void WorkWithListKSTeh(ref string error)
        {
            try
            {
                _forRecordWorkColumnInSmeta = new List<Dictionary<int, double>>();
                _nameAktKSRecordColumn = new List<string>();
              for(int numKS=0; numKS<_aktKSToOneSmeta.Count; numKS ++)
                {
                    string nameAktKS = null;
                    nameAktKS += ParserExcel.GetNameAktKS(_aktKSToOneSmeta[numKS]);
                    _forRecordWorkColumnInSmeta.Add(_aktKSinOrderSort[numKS].TotalScopeWorkAktKSone);
                    _nameAktKSRecordColumn.Add(nameAktKS);
                }
            }
            catch (COMException ex)
            {
                error += $"{ex.Message}\n";
            }
        }

        //метод записывает в файл копии сметы объемы из Актов КС-2, каждый месяц в новый столбец,
        //вставка столбцов идет за столбцом объемы по смете  
        private void RecordPartFileTehnadzor(int i, Smeta smeta, ref int nextInsertColumn)
        {
            ICollection<int> keyCollScopeWorkAktKSone = _forRecordWorkColumnInSmeta[i].Keys;
            InsertNextColumn(i, smeta, nextInsertColumn, keyCollScopeWorkAktKSone);
            Excel.Range topCellmergeCellNameAktKS = smeta.SheetDoc.Cells[smeta.RangeDoc.Row, nextInsertColumn];
            Excel.Range bottomCellmergeCellNameAktKS = smeta.SheetDoc.Cells[smeta.RangeDoc.Row + 2, nextInsertColumn];
            Excel.Range mergeCellNameAktKS = smeta.SheetDoc.get_Range(topCellmergeCellNameAktKS, bottomCellmergeCellNameAktKS);
            mergeCellNameAktKS.Merge();
            mergeCellNameAktKS.Value = _nameAktKSRecordColumn[i];
            smeta.SheetDoc.Cells[smeta.RangeDoc.Row + 3, nextInsertColumn] = nextInsertColumn - smeta.RangeDoc.Column + 1;
            nextInsertColumn += 1;
        }
        //метод записывает в файл копии сметы объемы из Актов КС-2 за каждый месяц 
        private void InsertNextColumn(int i, Smeta smeta, int nextInsertColumn, ICollection<int> keyCollScopeWorkAktKSone)
        {
            int pozSmeta;
            for (int j = smeta.RangeDoc.Row; j < smeta.RangeDoc.Rows.Count + smeta.RangeDoc.Row + 1; j++)
            {
                Excel.Range cellsNextColumnTablInsert = smeta.SheetDoc.Cells[j, nextInsertColumn];
                cellsNextColumnTablInsert.Insert(XlInsertShiftDirection.xlShiftToRight);
                if (j > smeta.RangeDoc.Row + 4)
                {
                    Excel.Range cellsNumPozColumnTabl = smeta.SheetDoc.Cells[j, smeta.RangeDoc.Column];
                    if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells)
                    {
                        try
                        {
                            pozSmeta = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                            foreach (int pozSmetaAktKS in keyCollScopeWorkAktKSone)
                            {
                                if (pozSmeta == pozSmetaAktKS)
                                {
                                    smeta.SheetDoc.Cells[j, nextInsertColumn] = _forRecordWorkColumnInSmeta[i][pozSmetaAktKS];
                                }
                            }
                        }
                        catch (FormatException ex)
                        {
                            smeta.Error += $"{ex.Message}\n";
                        }
                    }
                }
            }
        }


        //метод записывает в последний столбец "Остаток" формулу разности - остатка работ для технадзора
        private void RecordFormulaTehnadzor(Smeta smeta, int nextInsertColumn)
        {
            Excel.Range topInsertColumn = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row, nextInsertColumn];
            Excel.Range bottomInsertColumn = smeta.SheetDoc.Cells[smeta.RangeDoc.Rows.Count, nextInsertColumn];
            Excel.Range restInsertColumn = smeta.SheetDoc.get_Range(topInsertColumn, bottomInsertColumn);
            restInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topMergeCellContentRest = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row, nextInsertColumn];
            Excel.Range bottomMergeCellContentRest = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row + 2, nextInsertColumn];
            Excel.Range mergeCellContentRest = smeta.SheetDoc.get_Range(topMergeCellContentRest, bottomMergeCellContentRest);
            mergeCellContentRest.Merge();
            mergeCellContentRest.Value = "Остаток";
            mergeCellContentRest.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            mergeCellContentRest.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            mergeCellContentRest.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            mergeCellContentRest.EntireColumn.AutoFit();
            smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row + 3, nextInsertColumn] = nextInsertColumn - smeta.RangeDoc.Column + 1;
            Excel.Range cellContentNumRest = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row + 3, nextInsertColumn];
            cellContentNumRest.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            cellContentNumRest.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            cellContentNumRest.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            cellContentNumRest.EntireColumn.AutoFit();
            int amountColumnAktKS = nextInsertColumn - smeta.KeyConstructWorkSmeta.Column;
            if (amountColumnAktKS > 1)
            {
                for (int j = smeta.RangeDoc.Row + 4; j < smeta.RangeDoc.Rows.Count + smeta.RangeDoc.Row; j++)
                {
                    Excel.Range restFormula = smeta.SheetDoc.Cells[j, nextInsertColumn];
                    RecordMathFormula(smeta, restFormula, amountColumnAktKS, j);
                }
            }
        }
        //производит запись формулы
        private void RecordMathFormula(Smeta smeta, Excel.Range restFormula, int amountColumnAktKS, int j)
        {
            restFormula.Cells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            restFormula.EntireColumn.HorizontalAlignment = Excel.Constants.xlCenter;
            restFormula.EntireColumn.VerticalAlignment = Excel.Constants.xlCenter;
            restFormula.EntireColumn.AutoFit();
            Excel.Range cellsVupolnSmetaColumnTabl = smeta.SheetDoc.Cells[j, smeta.KeyConstructWorkSmeta.Column];
            if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
            {
                switch (amountColumnAktKS)
                {
                    case 2:
                        restFormula.FormulaR1C1 = "=RC[-2]-RC[-1]"; break;
                    case 3:
                        restFormula.FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"; break;
                    case 4:
                        restFormula.FormulaR1C1 = "=RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 5:
                        restFormula.FormulaR1C1 = "=RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 6:
                        restFormula.FormulaR1C1 = "=RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 7:
                        restFormula.FormulaR1C1 = "=RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 8:
                        restFormula.FormulaR1C1 = "=RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 9:
                        restFormula.FormulaR1C1 = "=RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 10:
                        restFormula.FormulaR1C1 = "=RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 11:
                        restFormula.FormulaR1C1 = "=RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 12:
                        restFormula.FormulaR1C1 = "=RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    case 13:
                        restFormula.FormulaR1C1 = "=RC[-13]-RC[-12]-RC[-11]-RC[-10]-RC[-9]-RC[-8]-RC[-7]-RC[-6]-RC[-5]-RC[-4]-RC[-3]-RC[-2]-RC[-1]"; break;
                    default: Console.WriteLine("Сводная таблица ведется до года, начните новую"); break;
                }
                restFormula.EntireColumn.AutoFit();
            }
        }
        //метод возвращает отсортированный лист книг Иксель, акты КС для сметы
        private void SortAktKSforTehnadzor()
        {
            Dictionary<string, int> numberList = new Dictionary<string, int>();
            for (int i = 0; i < _aktKSToOneSmeta.Count; i++)
            {
              FindForSortAktKS(i, ref numberList);
            }
            SortAktKS(numberList);           
        }
        //метод сортирует акты КС-2
        private void SortAktKS(Dictionary<string, int> numberList)
        {
            int[] valueNomerCifra = numberList.Values.ToArray();
            string[] keyNomerCifra = numberList.Keys.ToArray();
            SortDates(valueNomerCifra, keyNomerCifra);      
            for (int j = 0; j < keyNomerCifra.Length; j++)
            {
                for (int i = 0; i < _aktKSToOneSmeta.Count; i++)
                {
                    if (_aktKSToOneSmeta[i].DatAktKS.Contains(keyNomerCifra[j]))
                    {
                        _aktKSinOrderSort.Add(_aktKSToOneSmeta[i]); 
                        break;
                    }
                }
            }
        }
        //метод сортирует даты
        private void SortDates(int[] valueNomerCifra, string[]  keyNomerCifra)
        {
            for (int i = 1; i < valueNomerCifra.Length; i++)
            {
                for (int j = i; j > 0; j--)
                {
                    if (valueNomerCifra[j] < valueNomerCifra[j - 1])
                    {
                        int temp = valueNomerCifra[j - 1];
                        valueNomerCifra[j - 1] = valueNomerCifra[j];
                        valueNomerCifra[j] = temp;
                        string test = keyNomerCifra[j - 1];
                        keyNomerCifra[j - 1] = keyNomerCifra[j];
                        keyNomerCifra[j] = test;
                    }
                    else break;
                }
            }
        }
        //производит запись в словарь, где ключ - дата акта КС-2, значение - числовое выражение даты
        private void FindForSortAktKS(int num, ref Dictionary<string, int> numberList)
        {
                int number = _aktKSToOneSmeta[num].YearAktKS * 10000 + _aktKSToOneSmeta[num].MonthAktKS * 100 + _aktKSToOneSmeta[num].DayAktKS;
                numberList.Add(_aktKSToOneSmeta[num].DatAktKS, number);
        }
     

    }
}
