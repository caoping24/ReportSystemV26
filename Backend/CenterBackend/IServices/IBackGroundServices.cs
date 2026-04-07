using CenterReport.Repository.Models;

namespace CenterBackend.IServices
{
    public interface IBackGroundServices
    {
        Task Daily0810();
        Task WeeklyMon0820();
        Task MonthlyDay1_0830();
    }
}
