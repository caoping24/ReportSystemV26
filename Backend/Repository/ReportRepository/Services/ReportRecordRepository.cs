using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository.Services
{
    public class ReportRecordRepository<T> : IReportRecordRepository<T> where T : class
    {
        protected readonly CenterReportDbContext _context;
        private readonly DbSet<T> _entities;

        public ReportRecordRepository(CenterReportDbContext context)
        {
            _context = context;
            _entities = _context.Set<T>();

        }

        public IQueryable<T> db => _entities.AsQueryable();

        public async Task AddAsync(T entity)
        {
            await _context.AddAsync(entity);
        }

        public async Task<PaginationResult<ReportRecord>> GetReportByPageAsync(PaginationRequest request)
        {
            // 校验分页参数（避免无效参数）
            var pageIndex = Math.Max(1, request.PageIndex);
            var pageSize = Math.Clamp(request.PageSize, 1, 100); // 限制每页最大条数为100

            // 核心优化：先过滤，再排序（避免类型转换问题）
            var query = _context.ReportRecords
                               .AsNoTracking(); // 只读场景提升性能

            // 第一步：执行过滤条件（先过滤，不影响排序类型）
            query = query.Where(r => r.Type == request.Type);

            // 第二步：执行排序（此时 query 是 IQueryable，排序后转为 IOrderedQueryable）
            var orderedQuery = query.OrderByDescending(r => r.ReportedTime);

            // 第三步：基于排序后的查询执行分页
            var totalCount = await orderedQuery.LongCountAsync();
            var data = await orderedQuery
                .Skip((pageIndex - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            // 构造分页结果返回
            return new PaginationResult<ReportRecord>
            {
                PageIndex = pageIndex,
                PageSize = pageSize,
                TotalCount = totalCount,
                Data = data
            };
        }
    }
}
