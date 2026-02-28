using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using NPOI.HPSF;

namespace CenterBackend.Services
{
    public class DataToViewService(IReportRepository<SourceData> SourceData, IReportRepository<OperatorInputData> rperatorInputData)
    {

        private readonly IReportRepository<SourceData> _sourceData = SourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = rperatorInputData;

        public DayWorkBook DayGetMapData(DayWorkBook DayWorkBook, SourceData sourceData, OperatorInputData operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();
            var filteredSourceData = DayWorkBook.DaySheet;
            return DayWorkBook;

        }

        public MonthWorkBook MonthGetMapData(MonthWorkBook MonthWorkBook, SourceData sourceData, OperatorInputData operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return MonthWorkBook;

        }

        public YearWorkBook YearGetMapData(YearWorkBook YearWorkBook, SourceData sourceData, OperatorInputData operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();



            return YearWorkBook;

        }

        public WeekWorkBook WeekGetMapData(WeekWorkBook WeekWorkBook, SourceData sourceData, OperatorInputData operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return WeekWorkBook;

        }



    }
}
