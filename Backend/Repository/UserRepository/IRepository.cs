namespace CenterUser.Repository
{
    public interface IRepository<T> where T : class, ISoftDelete
    {
        IQueryable<T> db { get; }
        Task<T> GetByIdAsync(long id);
        Task AddAsync(T entity);
        void Update(T entity);
        Task DeleteByIdAsync(long id);
    }
}
