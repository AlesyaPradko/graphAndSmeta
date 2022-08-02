using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class Smeta:DocumentExcel
    {
        private readonly Excel.Range _keyNumberPosSmeta;
        private readonly Excel.Range _keyConstructWorkSmeta;
        
        public Excel.Range KeyNumberPosSmeta { get { return _keyNumberPosSmeta; } }
        public Excel.Range KeyConstructWorkSmeta { get { return _keyConstructWorkSmeta; } }
       
      
        public Smeta(string _name) :base(_name)
        {
            _keyNumberPosSmeta = FindText("№ пп", this, RangeDoc);
            _keyConstructWorkSmeta = FindText("Кол.", this, RangeDoc);      
        }
    }
} 