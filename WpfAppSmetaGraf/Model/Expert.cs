using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace WpfAppSmetaGraf.Model
{
    public enum XlInsertShiftDirection { xlShiftDown, xlShiftToRight };
    public class Expert : Worker
    {
        private Dictionary<int, double> _totalScopeWorkForSmeta;
        private Dictionary<int, string> _periodTimeWorkForSmeta;
        //выполняет работу по сбору информации и построению ведомости
        protected override void ProcessSmeta(int num, int size, ref string textError)
        {
            int lastRowCellsAfterDelete = 0;
            int insertColumnTotalScopeWork = _containCopySmeta[num].KeyConstructWorkSmeta.Column + 1;
            ParserExcel.DeleteColumnandRow(_containCopySmeta[num], ref lastRowCellsAfterDelete);
            Excel.Range newLastCell = _containCopySmeta[num].SheetDoc.Cells[lastRowCellsAfterDelete, _containCopySmeta[num].RangeDoc.Columns.Count];
            Excel.Range newFirstCell = _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].KeyNumberPosSmeta.Row, _containCopySmeta[num].KeyNumberPosSmeta.Column];
            _containCopySmeta[num].RangeDoc = _containCopySmeta[num].SheetDoc.get_Range(newFirstCell, newLastCell);//уменьшение области обработки
            InsertNewColumn(_containCopySmeta[num], insertColumnTotalScopeWork);
            _totalScopeWorkForSmeta = ParserExcel.GetkeySmetaForRecord<double>(_containCopySmeta[num]);
            _periodTimeWorkForSmeta = ParserExcel.GetkeySmetaForRecord<string>(_containCopySmeta[num]);
            Excel.Range CellContentConstruct = _containCopySmeta[num].RangeDoc.Find("Выполнение по смете");
            Excel.Range CellContentNote = _containCopySmeta[num].RangeDoc.Find("Примечание");
            int numberLastColumnCellNote;
            if (CellContentConstruct == null)
            {
                Implementation(num, insertColumnTotalScopeWork);
                numberLastColumnCellNote = GetColumforRecordNote(_containCopySmeta[num]);
            }
            else numberLastColumnCellNote = CellContentNote.Column;
            string error = null;
            if (_aktKSToOneSmeta.Count != 0)
            {
                WorkWithListKSExp(ref error);
            }
            _containCopySmeta[num].Error += error;
            RecordInFileExpert(_containCopySmeta[num], insertColumnTotalScopeWork, numberLastColumnCellNote);
            RecordFormulaExpert(_containCopySmeta[num], insertColumnTotalScopeWork);
            FormatRecordCopySmeta(_containCopySmeta[num], size);
            Excel.Range topLastColumnNote = _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].RangeDoc.Row, numberLastColumnCellNote + 1];
            Excel.Range bottomLastColumnNote = _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].RangeDoc.Rows.Count, numberLastColumnCellNote + 1];
            Excel.Range rangeLastColumnNote = _containCopySmeta[num].SheetDoc.get_Range(topLastColumnNote, bottomLastColumnNote);
            rangeLastColumnNote.ColumnWidth = 50;
            textError += _containCopySmeta[num].Error;
            ParserExcel.CloseDoc(_containCopySmeta[num]);
        }
        //создает столбец Выполнение по смете
        private void Implementation(int num, int insertColumnTotalScopeWork)
        {
            Excel.Range topMergeCellContentConstruct = _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].KeyNumberPosSmeta.Row, insertColumnTotalScopeWork];
            Excel.Range bottomMergeCellContentConstruct = _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].KeyNumberPosSmeta.Row + 2, insertColumnTotalScopeWork];
            Excel.Range mergeCellContentConstruct = _containCopySmeta[num].SheetDoc.get_Range(topMergeCellContentConstruct, bottomMergeCellContentConstruct);
            mergeCellContentConstruct.Merge();
            mergeCellContentConstruct.Value = "Выполнение по смете";
            _containCopySmeta[num].SheetDoc.Cells[_containCopySmeta[num].KeyNumberPosSmeta.Row + 3, insertColumnTotalScopeWork] = insertColumnTotalScopeWork - _containCopySmeta[num].KeyNumberPosSmeta.Column + 1;
        }
        //вставляет столбец
        private void InsertNewColumn(Smeta smeta, int insertColumnTotalScopeWork)
        {
            Excel.Range firstCellNewColumn = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, insertColumnTotalScopeWork];
            Excel.Range lastCellProcessing = smeta.SheetDoc.Range[RangeFile.LastCell];
            Excel.Range lastCellNewColumn = smeta.SheetDoc.Cells[lastCellProcessing.Row, insertColumnTotalScopeWork];
            Excel.Range insertNewColumn = smeta.SheetDoc.get_Range(firstCellNewColumn, lastCellNewColumn);
            insertNewColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
        }
        //создает столбец Примечание
        private int GetColumforRecordNote(Smeta smeta)
        {
            int numLastColumnCellNote = -1;
            for (int j = smeta.RangeDoc.Column; j <= smeta.RangeDoc.Columns.Count; j++)
            {
                Excel.Range cellsFirstRowTabl = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, j];
                if (cellsFirstRowTabl != null && cellsFirstRowTabl.Value2 != null || cellsFirstRowTabl.MergeCells) continue;
                else
                {
                    Excel.Range topCellmergeCellContentNote = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row, j];
                    Excel.Range bottomCellmergeCellContentNote = smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row + 2, j];
                    Excel.Range mergeCellContentNote = smeta.SheetDoc.get_Range(topCellmergeCellContentNote, bottomCellmergeCellContentNote);
                    mergeCellContentNote.Merge();
                    mergeCellContentNote.Value = "Примечание";
                    smeta.SheetDoc.Cells[smeta.KeyNumberPosSmeta.Row + 3, j] = j - smeta.RangeDoc.Column + 2;
                    numLastColumnCellNote = j;
                    break;
                }
            }
            return numLastColumnCellNote;
        }
        //производит работу с актами КС-2
        private void WorkWithListKSExp(ref string error)
        {
            string errorKS = null;
            try
            {
                string[] nameAktKS = new string[_aktKSToOneSmeta.Count];
                Parallel.For(0, _aktKSToOneSmeta.Count, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount / 2 }, numKS =>
                {
                    nameAktKS[numKS] = ParserExcel.GetNameAktKS(_aktKSToOneSmeta[numKS]);
                    int[] keyScopeWorkforSmeta = _totalScopeWorkForSmeta.Keys.ToArray();
                    int[] keyWorkAktKS = _aktKSToOneSmeta[numKS].TotalScopeWorkAktKSone.Keys.ToArray();
                    for (int i = 0; i < _totalScopeWorkForSmeta.Count; i++)
                    {
                        for (int j = 0; j < keyWorkAktKS.Length; j++)
                        {
                            if (keyScopeWorkforSmeta[i] == keyWorkAktKS[j])
                            {
                                _totalScopeWorkForSmeta[keyWorkAktKS[j]] += _aktKSToOneSmeta[numKS].TotalScopeWorkAktKSone[keyWorkAktKS[j]];
                                _periodTimeWorkForSmeta[keyWorkAktKS[j]] += nameAktKS[numKS];
                            }
                        }
                    }
                    errorKS += _aktKSToOneSmeta[numKS].Error;
                });
            }
            catch (COMException ex)
            {
                errorKS += $"{ex.Message}\n";
            }
            ParserExcel.CloseAktKS(_aktKSToOneSmeta);
            error = errorKS;
        }
        //запись данных в ведомость
        private void RecordInFileExpert(Smeta smeta, int insertColumnTotalScopeWork, int numberLastColumnCellNote)
        {
            int pozSmeta;
            for (int j = smeta.RangeDoc.Row + 4; j < smeta.RangeDoc.Rows.Count + smeta.RangeDoc.Row + 4; j++)
            {
                Excel.Range cellsNumPozColumnTabl = smeta.SheetDoc.Cells[j, smeta.RangeDoc.Column];
                if (cellsNumPozColumnTabl != null && cellsNumPozColumnTabl.Value2 != null && cellsNumPozColumnTabl.Value2.ToString() != "" && !cellsNumPozColumnTabl.MergeCells)
                {
                    try
                    {
                        pozSmeta = Convert.ToInt32(cellsNumPozColumnTabl.Value2);
                        smeta.SheetDoc.Cells[j, insertColumnTotalScopeWork] = _totalScopeWorkForSmeta[pozSmeta];
                        smeta.SheetDoc.Cells[j, numberLastColumnCellNote] = _periodTimeWorkForSmeta[pozSmeta];
                    }
                    catch (FormatException ex)
                    {
                        smeta.Error += $"{ex.Message}\n";
                    }
                }
            }
        }
        //производит запись формулы 
        private void RecordFormulaExpert(Smeta smeta, int vstavkaColumntotalScopeWork)
        {
            Excel.Range topInsertColumn = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomInsertColumn = smeta.SheetDoc.Cells[smeta.RangeDoc.Rows.Count, vstavkaColumntotalScopeWork + 1];
            Excel.Range restInsertColumn = smeta.SheetDoc.get_Range(topInsertColumn, bottomInsertColumn);
            restInsertColumn.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
            Excel.Range topMergeCellContentRest = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row, vstavkaColumntotalScopeWork + 1];
            Excel.Range bottomMergeCellContentRest = smeta.SheetDoc.Cells[smeta.KeyConstructWorkSmeta.Row + 2, vstavkaColumntotalScopeWork + 1];
            Excel.Range mergeCellContentRest = smeta.SheetDoc.get_Range(topMergeCellContentRest, bottomMergeCellContentRest);
            mergeCellContentRest.Merge();
            mergeCellContentRest.Value = "Остаток";
            smeta.SheetDoc.Cells[smeta.RangeDoc.Row + 3, vstavkaColumntotalScopeWork + 1] = vstavkaColumntotalScopeWork - smeta.RangeDoc.Column + 2;
            for (int j = smeta.RangeDoc.Row + 4; j < smeta.RangeDoc.Rows.Count + smeta.RangeDoc.Row + 4; j++)
            {
                Excel.Range cellsVupolnSmetaColumnTabl = smeta.SheetDoc.Cells[j, smeta.KeyConstructWorkSmeta.Column];
                if (cellsVupolnSmetaColumnTabl != null && cellsVupolnSmetaColumnTabl.Value2 != null && cellsVupolnSmetaColumnTabl.Value2.ToString() != "" && !cellsVupolnSmetaColumnTabl.MergeCells)
                {
                    Excel.Range restFormula = smeta.SheetDoc.Cells[j, vstavkaColumntotalScopeWork + 1];
                    restFormula.FormulaR1C1 = "=RC[-2]-RC[-1]";
                }
            }
        }

    }
}
