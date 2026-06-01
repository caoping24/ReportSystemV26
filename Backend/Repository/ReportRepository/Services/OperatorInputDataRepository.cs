using CenterReport.Repository.IServices;
using Microsoft.EntityFrameworkCore;

namespace CenterReport.Repository.Services
{
    public class OperatorInputDataRepository<T> : IOperatorInputDataRepository<T> where T : class
    {
        protected readonly CenterReportDbContext _context;
        private readonly DbSet<T> _entities;

        public OperatorInputDataRepository(CenterReportDbContext context)
        {
            _context = context;
            _entities = _context.Set<T>();
        }

        public IQueryable<T> Db => _entities.AsQueryable();
        public async Task<T?> GetByIdAsync(long id) => await _entities.FindAsync(id);
        public async Task AddAsync(T entity) => await _entities.AddAsync(entity);
        public async Task Update(T entity) => _context.Entry(entity).State = EntityState.Modified;
        public async Task<T> UpsertByIdAsync(T entity, Action<T> updateAction)
        {
            // 校验入参
            ArgumentNullException.ThrowIfNull(entity, nameof(entity));
            ArgumentNullException.ThrowIfNull(updateAction, nameof(updateAction));

            var idValue = EF.Property<long>(entity, "Id");// 从实体中提取ID值
            T? existingEntity;
            if (idValue == 0)
            {
                await _entities.AddAsync(entity);
                existingEntity = entity;
            }
            else
            {
                existingEntity = await _entities.FindAsync(idValue);
                if (existingEntity != null)
                {
                    updateAction(existingEntity);//存在则执行自定义更新逻辑
                }

            }
            return existingEntity;
        }

    }
}
