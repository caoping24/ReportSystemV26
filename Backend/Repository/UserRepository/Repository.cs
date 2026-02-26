using Microsoft.EntityFrameworkCore;


namespace CenterUser.Repository
{
    public class Repository<T> : IRepository<T> where T : class, ISoftDelete
    {
        protected readonly ApplicationDbContext _context;
        private DbSet<T> _entities;

        public Repository(ApplicationDbContext context)
        {
            _context = context;
            _entities = _context.Set<T>();
        }

        public IQueryable<T> db => _entities.AsQueryable();

        public async Task<T> GetByIdAsync(long id) => await _entities.FindAsync(id);

        public async Task AddAsync(T entity)
        {
            await _entities.AddAsync(entity);
        }

        public void Update(T entity)
        {
            _context.Entry(entity).State = EntityState.Modified;
        }

        public async Task DeleteByIdAsync(long id)
        {
            var entity = await _context.Set<T>().FindAsync(id);
            if (entity != null)
            {
                entity.IsDelete = true;
                _context.Entry(entity).State = EntityState.Modified;
            }
        }
    }
}
