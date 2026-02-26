using CenterReport.Repository.Models;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository
{
    public class CenterReportDbContext : DbContext
    {

        public DbSet<SourceData> SourceDatas => Set<SourceData>();
        public DbSet<OperatorInputData> OperatorInputDatas => Set<OperatorInputData>();
        public DbSet<CalculatedData> CalculatedDatas => Set<CalculatedData>();
        public DbSet<ReportRecord> ReportRecords => Set<ReportRecord>();
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
