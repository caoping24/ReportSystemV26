using CenterBackend.Models.CalculateData;

namespace CenterBackend.IServices
{
    public interface ICalculatedAndSaveService
    {
        Task<bool> DataAnalyses(ReportInfo ReportInfo);
    }
}
