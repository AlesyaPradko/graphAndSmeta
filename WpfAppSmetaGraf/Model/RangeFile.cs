

namespace WpfAppSmetaGraf.Model
{
    public static class RangeFile
    {
        private readonly static string _firstCell= "A1";
        private readonly static string _lastCell= "AD2200";
        public static string FirstCell { get { return _firstCell; } }
        public static string LastCell { get { return _lastCell;} }
    }
}