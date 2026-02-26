using CenterReport.Repository.Models;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository
{
    public class CenterReportDbContext : DbContext
    {

        public DbSet<SourceData> SourceData => Set<SourceData>();

        public DbSet<CalculatedData> CalculatedDatas => Set<CalculatedData>();

        public DbSet<ReportRecord> ReportRecord => Set<ReportRecord>();
        public CenterReportDbContext(DbContextOptions<CenterReportDbContext> options)
               : base(options)
        {
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
        }
    }
}
