using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;


namespace CenterReport.Repository.IServices
{
    public interface IReportRecordRepository<T> where T : class
    {
        Task<PaginationResult<ReportRecord>> GetReportByPageAsync(PaginationRequest request);
        IQueryable<T> Db { get; }
        Task<T?> GetByIdAsync(long id);
        Task AddAsync(T entity);
        Task Update(T entity);
        Task<T> UpsertByIdAsync(T entity, Action<T> updateAction);
    }
}
