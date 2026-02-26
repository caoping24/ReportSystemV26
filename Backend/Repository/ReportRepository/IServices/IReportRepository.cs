using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository.IServices
{
    public interface IReportRepository<T> where T : class
    {
        Task<List<T>> GetByDateTimeRangeAsync(DateTime startTime, DateTime endTime); 
        Task<List<T>> GetByDateTimeRangeAsync(DateTime startTime, DateTime endTime, int dataType);
        Task<List<T>> GetByExactDateTime(DateTime targetTime, int dataType);
        IQueryable<T> Db { get; }
        Task<T?> GetByIdAsync(long id);
        Task AddAsync(T entity);
        Task Update(T entity);
        Task DeleteByIdAsync(long id);
    }
}
