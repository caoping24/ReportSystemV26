using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;


namespace CenterReport.Repository
{
    public interface IReportRecordRepository<T> where T : class
    {
        IQueryable<T> db { get; }
        Task<PaginationResult<ReportRecord>> GetReportByPageAsync(PaginationRequest request);
        Task AddAsync(T entity);
    }
}
