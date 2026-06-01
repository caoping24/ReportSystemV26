using CenterBackend.Models;

namespace CenterBackend.IServices
{
    public interface IReportService
    {
        Task<bool> RebuildReport(PathAndName fileInfo);
    }
}
