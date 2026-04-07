using CenterReport.Repository.Models;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository
{
    public class CenterReportDbContext : DbContext
    {

        public DbSet<SourceData> SourceDatas => Set<SourceData>();
        public DbSet<OperatorInputData> OperatorInputDatas => Set<OperatorInputData>();
        public DbSet<ReportRecord> ReportRecords => Set<ReportRecord>();
        public DbSet<ComToSiemens> ComToSiemenss => Set<ComToSiemens>(); //2026年4月22日新增

        public CenterReportDbContext(DbContextOptions<CenterReportDbContext> options)
               : base(options)
        {
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<OperatorInputData>()
       .ToTable(tb => tb.HasTrigger("TR_OperatorInputData_UpdateLastChange"));
      
        }
    }
}
