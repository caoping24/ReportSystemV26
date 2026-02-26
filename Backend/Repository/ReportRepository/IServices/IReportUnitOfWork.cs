namespace CenterReport.Repository.IServices
{
    public interface IReportUnitOfWork
    {
        Task<int> SaveChangesAsync();
    }
}
