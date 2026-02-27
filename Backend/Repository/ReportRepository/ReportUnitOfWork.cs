namespace CenterReport.Repository
{
    public class ReportUnitOfWork(CenterReportDbContext context) : IReportUnitOfWork
    {
        private readonly CenterReportDbContext _context = context;

        public async Task<int> SaveChangesAsync()
        {
            return await _context.SaveChangesAsync();
        }
    }
}
