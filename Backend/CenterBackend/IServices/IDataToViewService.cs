using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;

namespace CenterBackend.IServices
{
    public interface IDataToViewService
    {
        Task<bool> DayGetMapDataAsync(DayWorkBook DayWorkBook);
        Task<bool> MonthGetMapDataAsync(MonthWorkBook monthWorkBook);
        Task<bool> YearGetMapDataAsync(YearWorkBook yearWorkBook);
        Task<bool> WeekGetMapDataAsync(WeekWorkBook WeekWorkBook);
    }
}
