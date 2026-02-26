using CenterReport.Repository.IServices;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository.Services
{
    public class ReportRepository<T> : IReportRepository<T> where T : class
    {
        protected readonly CenterReportDbContext _context;
        private DbSet<T> _entities;

        public ReportRepository(CenterReportDbContext context)
        {
            _context = context;
            _entities = _context.Set<T>();
        }

        // 根据 时间范围 查询记录
        public async Task<List<T>> GetByDateTimeRangeAsync(DateTime startTime, DateTime endTime)
        {
            var actualStartTime = startTime < endTime ? startTime : endTime;
            var actualEndTime = startTime < endTime ? endTime : startTime;

            return await _entities
                .Where(e =>
                    EF.Property<DateTime>(e, "ReportedTime") >= actualStartTime &&
                    EF.Property<DateTime>(e, "ReportedTime") <= actualEndTime)
                .OrderBy(e => EF.Property<DateTime>(e, "ReportedTime"))
                .ToListAsync();
        }

        // 根据 时间范围+数据类型 查询记录(重载)
        public async Task<List<T>> GetByDateTimeRangeAsync(DateTime startTime, DateTime endTime, int dataType)
        {
            var startDate = startTime < endTime ? startTime : endTime;
            var endDate = startTime < endTime ? endTime : startTime;

            return startDate == endDate
                ? await GetByExactDateTime(startDate, dataType)
                : await QueryByDateTimeRangeAndType(startDate, endDate, dataType);
        }

        // 根据 精确时间+数据类型 查询记录
        public async Task<List<T>> GetByExactDateTime(DateTime targetTime, int dataType)
        {
            return await _entities
            .Where(e =>
                    EF.Property<DateTime>(e, "ReportedTime") == targetTime &&
                    EF.Property<int>(e, "Type") == dataType)
            .OrderBy(e => EF.Property<DateTime>(e, "ReportedTime"))// 统一排序规则
            .ToListAsync();
        }
        public IQueryable<T> Db => _entities.AsQueryable();
        public async Task<T?> GetByIdAsync(long id) => await _entities.FindAsync(id);
        public async Task AddAsync(T entity) => await _entities.AddAsync(entity);
        public async Task Update(T entity) => _context.Entry(entity).State = EntityState.Modified;
        public async Task DeleteByIdAsync(long id)
        {
            var entity = await _context.Set<T>().FindAsync(id);
            if (entity != null)
            {
                _context.Remove(entity);
            }
        }

        /// <summary>
        /// 私有方法
        /// </summary>
        //时间范围+数据类型查询
        private async Task<List<T>> QueryByDateTimeRangeAndType(DateTime startDate, DateTime endDate, int dataType)
        {
            return await _entities
                .Where(e =>
                    EF.Property<DateTime>(e, "ReportedTime") >= startDate &&
                    EF.Property<DateTime>(e, "ReportedTime") <= endDate &&
                    EF.Property<int>(e, "Type") == dataType)
                .OrderBy(e => EF.Property<DateTime>(e, "ReportedTime"))
                .ToListAsync();
        }


    }
}
