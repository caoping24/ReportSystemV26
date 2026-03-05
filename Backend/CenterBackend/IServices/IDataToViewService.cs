using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;

namespace CenterBackend.IServices
{
    public interface IDataToViewService
    {

        Task<bool> DayGetMapDataAsync(DayWorkBook DayWorkBook);
        Task<bool> MonthGetMapDataAsync(WeekWorkBook WeekWorkBook);
        Task<bool> YearGetMapDataAsync(WeekWorkBook WeekWorkBook);
        Task<bool> WeekGetMapDataAsync(WeekWorkBook WeekWorkBook);

    }
}