using Microsoft.EntityFrameworkCore;
using Microsoft.VisualBasic;
using System;

namespace CenterReport.Repository
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

        public IQueryable<T> db => _entities.AsQueryable();

        /// <summary>
        /// 查询一天内的数据（昨日08:00-今日08:00）日统计专用
        /// </summary>
        /// <param name="time">日期时间,自动切片</param>
        /// <returns></returns>
        public async Task<List<T>> GetByDayAsync(DateTime time)
        {

            DateTime startDate = time.Date.AddHours(8);//今日08:00
            DateTime endDate = startDate.AddHours(24).AddMinutes(59);//明日08:00
            // 3. 执行查询：筛选当天数据 + 正序排序
            return await _entities
                .Where(e =>
                    EF.Property<DateTime>(e, "createdtime") >= startDate// 匹配当天所有时间（忽略createdtime的时分秒）
                    && EF.Property<DateTime>(e, "createdtime") < endDate)
                .OrderBy(e => EF.Property<DateTime>(e, "createdtime")) // 正序排序（默认ASC）
                .ToListAsync();
        }
        public async Task<List<T>> GetByDayAsyncType(DateTime queryDate, int? type = null)
        {
            DateTime startTime = queryDate.Date.AddHours(8);
            DateTime endTime = startTime.Date.AddHours(33);

            var query = _entities.Where(e =>
                    EF.Property<DateTime>(e, "createdtime") >= startTime// 匹配当天所有时间（忽略createdtime的时分秒）
                    && EF.Property<DateTime>(e, "createdtime") < endTime);

            // 可选按Type筛选
            if (type.HasValue)
            {
                query = query.Where(e => EF.Property<int>(e, "Type") == type.Value);
            }

            return await query.OrderBy(e => EF.Property<DateTime>(e, "createdtime")).ToListAsync();
        }

        /// <summary>
        /// SourceData表使用，该表没有reportedTime字段
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public async Task<List<T>> GetByDataTimeAsync(DateTime start, DateTime end)
        {
            var fromDateTime = DateTime.Compare(start, end) > 0 ? end : start;
            var toDateTime = DateTime.Compare(start, end) > 0 ? start : end;

            return await _entities
                .Where(e => EF.Property<DateTime>(e, "createdtime") >= fromDateTime && EF.Property<DateTime>(e, "createdtime") <= toDateTime)
                .OrderBy(e => EF.Property<DateTime>(e, "createdtime"))
                .ToListAsync();
        }
        /// <summary>
        /// CalculatedData表使用，该表有reportedTime字段
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public async Task<List<T>> GetByDataTimeAsync(DateTime start, DateTime end, int type)
        {
            DateTime fromDateTime = DateTime.Compare(start, end) > 0 ? end : start;
            DateTime toDateTime = DateTime.Compare(start, end) > 0 ? start : end;

            if (fromDateTime == toDateTime)
            {
                return await _entities
                    .Where(e =>
                        EF.Property<DateTime>(e, "reportedTime") == fromDateTime &&
                        EF.Property<int>(e, "Type") == type)
                    .OrderBy(e => EF.Property<DateTime>(e, "reportedTime"))// 统一排序规则
                    .ToListAsync();
            }
            return await _entities
                .Where(e =>
                    EF.Property<DateTime?>(e, "reportedTime") != null &&
                    EF.Property<DateTime>(e, "reportedTime") >= fromDateTime &&
                    EF.Property<DateTime>(e, "reportedTime") <= toDateTime &&
                    EF.Property<int>(e, "Type") == type)
                .OrderBy(e => EF.Property<DateTime>(e, "reportedTime"))
                .ToListAsync();
        }

        public async Task<T?> GetByIdAsync(int id) => await _entities.FindAsync(id);
        public async Task AddAsync(T entity)
        {
            await _entities.AddAsync(entity);
        }

        public async Task Update(T entity)
        {
            _context.Entry(entity).State = EntityState.Modified;
        }

        public async Task DeleteByIdAsync(int id)
        {
            var entity = await _context.Set<T>().FindAsync(id);
            if (entity != null)
            {
                _context.Remove(entity);
            }
        }


    }
}
