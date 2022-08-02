
using System.Collections.Generic;


namespace WpfAppSmetaGraf.Model
{
    public class Estimate
    {
        public int Id { get; set; }
        public string EstimateName { get; set; }
        public ICollection<Chapter> Chapters { get; set; }
    }
}
