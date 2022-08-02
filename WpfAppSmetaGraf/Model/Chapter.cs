

namespace WpfAppSmetaGraf.Model
{
    public class Chapter
    {
        public int Id { get; set; }
        public string ChapterName { get; set; }
        public string WorkName { get; set; }
        public int? EstimateId { get; set; }
        public Estimate Estimate { get; set; }
    }
}
