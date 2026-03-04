using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;

namespace CenterBackend.IServices
{
    public interface IDataToViewService
    {

        Task<bool> DayGetMapDataAsync(DayWorkBook DayWorkBook);
        bool MonthGetMapData(MonthWorkBook MonthWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData);
        bool YearGetMapData(YearWorkBook YearWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData);
        bool WeekGetMapData(WeekWorkBook WeekWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData);
    }
}