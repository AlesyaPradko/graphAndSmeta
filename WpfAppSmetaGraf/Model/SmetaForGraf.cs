using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class SmetaForGraf : Smeta
    {
        private readonly Excel.Range _keyTrudozatratSmeta;
        private readonly Excel.Range _cellWithTrudozatrat;
        private readonly Excel.Range _keyCellNameWork;
        private Dictionary<Excel.Range, double> _onChapterTrudozatrat;
        private readonly double _trudozatratTotal;
        private List<Excel.Range> _cellsAllChapter;
        private List<int> _startChapter;
        public Dictionary<Excel.Range, double> OnChapterTrudozatrat { get { return _onChapterTrudozatrat; } }
        public double TrudozatratTotal { get { return _trudozatratTotal; } }
        public Excel.Range KeyTrudozatratSmeta { get { return _keyTrudozatratSmeta; } }
        public Excel.Range CellWithTrudozatrat { get { return _cellWithTrudozatrat; } }
        public Excel.Range KeyCellNameWork { get { return _keyCellNameWork; } }
        public List<Excel.Range> CellsAllChapter { get { return _cellsAllChapter; } set { _cellsAllChapter = value; } }
       
        public List<int> StartChapter { get { return _startChapter; } set { _startChapter = value; } }
        public SmetaForGraf(string _name) : base(_name)
        {
            _keyTrudozatratSmeta = FindText("Т/з осн. раб. Всего",this, RangeDoc);
            _cellWithTrudozatrat = FindText("Сметная трудоемкость", this, RangeDoc);
            _keyCellNameWork = FindNameWork();
            _cellsAllChapter = ParserExcel.FindChapter(this);
            _trudozatratTotal = ParserExcel.NumeralFromCell(_cellWithTrudozatrat.Value.ToString());
            _onChapterTrudozatrat = ParserExcel.FindForChapter(this);
            _startChapter = GetFirstPosChapter();
        }
        //возвращает ячейку с содержимым Наименование
        private Excel.Range FindNameWork()
        {
            Excel.Range rangeSmetaForCell1 = SheetDoc.Cells[KeyNumberPosSmeta.Row, 1];
            Excel.Range rangeSmetaForCell2 = SheetDoc.Cells[KeyNumberPosSmeta.Row + 3, 10];
            Excel.Range rangeSmetaForCell = SheetDoc.get_Range(rangeSmetaForCell1, rangeSmetaForCell2);
            Excel.Range keyCellNameWork = FindText("Наименование", this, rangeSmetaForCell);
            return keyCellNameWork;
        }
        //возвращает лист с номераит позиций разделов
        private List<int> GetFirstPosChapter()
        {
            List<int> startChapter = new List<int>();
            try
            {
                for (int j = 0; j < _cellsAllChapter.Count; j++)
                {
                    Excel.Range startChapt = SheetDoc.Cells[_cellsAllChapter[j].Row + 1, _cellsAllChapter[j].Column];
                    if (startChapt != null && startChapt.Value2 != null && !startChapt.MergeCells && startChapt.Value2.ToString() != "" && startChapt != null)
                    {
                        startChapter.Add(Convert.ToInt32(startChapt.Value2));
                    }
                    else
                    {
                        startChapt = SheetDoc.Cells[_cellsAllChapter[j].Row + 2, _cellsAllChapter[j].Column];
                        if (startChapt != null && startChapt.Value2 != null && !startChapt.MergeCells && startChapt.Value2.ToString() != "" && startChapt != null)
                        {
                            startChapter.Add(Convert.ToInt32(startChapt.Value2));
                        }
                    }
                }
            }
            catch (FormatException exc)
            {
                Error += $"{exc.Message} Проверьте первый столбец и первые строки после разделов\n";
            }
            return startChapter;
        }
    }
}
