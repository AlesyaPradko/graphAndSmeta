using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfAppSmetaGraf.Model
{
    public class CheckIt
    {
        private readonly static Excel.Application instance = new Excel.Application();
        public static Excel.Application Instance
        {
            get
            {
                if (instance == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return null;
                }
                return instance;
            }
        }
        static CheckIt()
        { }
        private CheckIt()
        { }
    }
}
