using CenterReport.Repository.IServices;

namespace CenterReport.Repository.Services
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
