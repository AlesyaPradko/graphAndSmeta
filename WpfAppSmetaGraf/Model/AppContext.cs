
using System.Data.Entity;

namespace WpfAppSmetaGraf.Model
{
    class AppContext : DbContext
    {
        public AppContext() : base("DbConnection")
        { }
        public DbSet<Chapter> Chapters { get; set; }
        public DbSet<Estimate> Estimates { get; set; }
        public DbSet<WorkingDay> WorkingDays { get; set; }
    }
}
