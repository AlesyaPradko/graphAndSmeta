using System.Text.RegularExpressions;

namespace WpfAppSmetaGraf.Model
{
    public static class RegexReg
    {
        public static Regex ScopeWorkInAktKS { get { return new Regex(@"((К|к)оличество|Кол\.)", RegexOptions.IgnoreCase); } }
        public static Regex RegexDay { get { return new Regex(@"(?<day>\d{2})\.", RegexOptions.IgnoreCase); } }
        public static Regex RegexMonth { get { return new Regex(@"\.?(?<month>\d{2})\.", RegexOptions.IgnoreCase); } }
        public static Regex RegexYear { get { return new Regex(@"\.(?<year>\d{4})", RegexOptions.IgnoreCase); } }
        public static Regex RegexData { get { return new Regex(@"(?<month>\d{2})\.(?<year>\d{4})", RegexOptions.IgnoreCase); } }
        public static Regex RegexAllData { get { return new Regex(@"(?<day>\d{2})\.(?<month>\d{2})\.(?<year>\d{4})", RegexOptions.IgnoreCase); } }
        public static Regex NameSmeta { get { return new Regex(@"((С|с)мета|\s*) №\s*\d+", RegexOptions.IgnoreCase); } }
        public static Regex CellTotalForChapter { get { return new Regex("Итого по разделу"); } }
        public static Regex CellOfRazdel { get { return new Regex(@"^Раздел"); } }
    }
}