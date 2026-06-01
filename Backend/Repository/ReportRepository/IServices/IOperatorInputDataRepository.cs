namespace CenterReport.Repository.IServices
{
    public interface IOperatorInputDataRepository<T> where T : class
    {
        IQueryable<T> Db { get; }
        Task<T?> GetByIdAsync(long id);
        Task AddAsync(T entity);
        Task Update(T entity);
        Task<T> UpsertByIdAsync(T entity, Action<T> updateAction);
    }
}
