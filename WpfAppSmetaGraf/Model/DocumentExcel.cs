
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class DocumentExcel
    {

        protected string _addressDoc;
        protected Excel.Workbook _doc;
        protected Excel.Worksheet _sheetDoc;
        protected Excel.Range _rangeDoc;
        private string _error;
        public string Error { set { _error = value; } get { return _error; } }
        public string AddressDoc { get { return _addressDoc; } }
        public Excel.Workbook DocCur { set { _doc = value; } get { return _doc; } }
        public Excel.Worksheet SheetDoc { set { _sheetDoc = value; } get { return _sheetDoc; } }
        public Excel.Range RangeDoc { set { _rangeDoc = value; } get { return _rangeDoc; } }
        public DocumentExcel(string _name)
        {
            _addressDoc = _name;
            _doc = CheckIt.Instance.Workbooks.Open(_name);
            _sheetDoc = _doc.Sheets[1];
            _rangeDoc = _sheetDoc.get_Range(RangeFile.FirstCell, RangeFile.LastCell);
        }
        //находит требуемый текст
        protected Excel.Range FindText(string str, DocumentExcel doc, Excel.Range range)
        {
            Excel.Range key = range.Find(str);
            if (key == null)
            {
                string er = $" Проверьте чтобы в {AddressDoc} было верно записано устойчивое выражение {str}\n";
                Error += er;
                ParserExcel.CloseDoc(doc);
                throw new NullValueException(er);
            }
            return key;
        }
    }
}
