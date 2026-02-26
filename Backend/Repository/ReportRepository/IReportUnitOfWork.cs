namespace CenterReport.Repository
{
    public interface IReportUnitOfWork
    {
        Task<int> SaveChangesAsync();
    }
}
