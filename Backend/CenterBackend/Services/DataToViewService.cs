using CenterBackend.IServices;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using NPOI.HPSF;
using NPOI.SS.UserModel;

namespace CenterBackend.Services
{
    public class DataToViewService(IReportRepository<SourceData> SourceData, IReportRepository<OperatorInputData> rperatorInputData) : IDataToViewService
    {

        private readonly IReportRepository<SourceData> _sourceData = SourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = rperatorInputData;

        public bool DayGetMapData(DayWorkBook DayWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            var startTime = DayWorkBook.ReportedTime.Date;
            DayWorkBook.DaySheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();
            DayWorkBook.NightSheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();

            if (sourceData == null || sourceData.Count == 0)
                return false;
            var baseDate = startTime.AddHours(8);
            var dataPart1 = SortDataByTime(sourceData, baseDate, 25);
            var dataPart2 = SortDataByTime(operatorInputData, baseDate, 25);
            for (int i = 0; i < 13; i++)//白班
            {

                if (dataPart1 == null || dataPart1[i] == null)
                {
                    continue;
                }
                DayWorkBook.DaySheet[i].Cell1 = dataPart1[i].Cell1;
                DayWorkBook.DaySheet[i].Cell2 = dataPart1[i].Cell2;
                DayWorkBook.DaySheet[i].Cell3 = dataPart1[i].Cell3;
                DayWorkBook.DaySheet[i].Cell4 = dataPart1[i].Cell4;
                DayWorkBook.DaySheet[i].Cell5 = dataPart1[i].Cell5;
                DayWorkBook.DaySheet[i].Cell6 = dataPart1[i].Cell6;
                DayWorkBook.DaySheet[i].Cell7 = dataPart1[i].Cell7;
                DayWorkBook.DaySheet[i].Cell8 = dataPart1[i].Cell8;
                DayWorkBook.DaySheet[i].Cell9 = dataPart1[i].Cell9;
                DayWorkBook.DaySheet[i].Cell10 = dataPart1[i].Cell10;

                if (dataPart2 == null || dataPart2[i] == null)
                {
                    continue;
                }
                DayWorkBook.DaySheet[i].Cell1 = dataPart2[i].Cell1;
                DayWorkBook.DaySheet[i].Cell2 = dataPart2[i].Cell2;
                DayWorkBook.DaySheet[i].Cell3 = dataPart2[i].Cell3;
                DayWorkBook.DaySheet[i].Cell4 = dataPart2[i].Cell4;
                DayWorkBook.DaySheet[i].Cell5 = dataPart2[i].Cell5;
                DayWorkBook.DaySheet[i].Cell6 = dataPart2[i].Cell6;
                DayWorkBook.DaySheet[i].Cell7 = dataPart2[i].Cell7;
                DayWorkBook.DaySheet[i].Cell8 = dataPart2[i].Cell8;
                DayWorkBook.DaySheet[i].Cell9 = dataPart2[i].Cell9;
                DayWorkBook.DaySheet[i].Cell10 = dataPart2[i].Cell10;
            }

            for (int i = 0; i < 13; i++)//夜班
            {

                if (dataPart1 == null || dataPart1[i + 12] == null)
                {
                    continue;
                }
                DayWorkBook.NightSheet[i].Cell1 = dataPart1[i + 12].Cell1;
                DayWorkBook.NightSheet[i].Cell2 = dataPart1[i + 12].Cell2;
                DayWorkBook.NightSheet[i].Cell3 = dataPart1[i + 12].Cell3;
                DayWorkBook.NightSheet[i].Cell4 = dataPart1[i + 12].Cell4;
                DayWorkBook.NightSheet[i].Cell5 = dataPart1[i + 12].Cell5;
                DayWorkBook.NightSheet[i].Cell6 = dataPart1[i + 12].Cell6;
                DayWorkBook.NightSheet[i].Cell7 = dataPart1[i + 12].Cell7;
                DayWorkBook.NightSheet[i].Cell8 = dataPart1[i + 12].Cell8;
                DayWorkBook.NightSheet[i].Cell9 = dataPart1[i + 12].Cell9;
                DayWorkBook.NightSheet[i].Cell10 = dataPart1[i + 12].Cell10;

                if (dataPart2 == null || dataPart2[i + 12] == null)
                {
                    continue;
                }
                DayWorkBook.DaySheet[i].Cell1 = dataPart2[i + 12].Cell1;
                DayWorkBook.DaySheet[i].Cell2 = dataPart2[i + 12].Cell2;
                DayWorkBook.DaySheet[i].Cell3 = dataPart2[i + 12].Cell3;
                DayWorkBook.DaySheet[i].Cell4 = dataPart2[i + 12].Cell4;
                DayWorkBook.DaySheet[i].Cell5 = dataPart2[i + 12].Cell5;
                DayWorkBook.DaySheet[i].Cell6 = dataPart2[i + 12].Cell6;
                DayWorkBook.DaySheet[i].Cell7 = dataPart2[i + 12].Cell7;
                DayWorkBook.DaySheet[i].Cell8 = dataPart2[i + 12].Cell8;
                DayWorkBook.DaySheet[i].Cell9 = dataPart2[i + 12].Cell9;
                DayWorkBook.DaySheet[i].Cell10 = dataPart2[i + 12].Cell10;
            }
 
            return true;

        }

        public bool MonthGetMapData(MonthWorkBook MonthWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return false;

        }

        public bool YearGetMapData(YearWorkBook YearWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();



            return false;

        }

        public bool WeekGetMapData(WeekWorkBook WeekWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return false;

        }

        //根据时间排序数据-原始数据
        private static List<SourceData> SortDataByTime(List<SourceData> sourceData, DateTime baseDate, int maxCount)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<SourceData>(new SourceData[maxCount]);
            for (int i = 0; i < maxCount; i++)
            {
                DateTime intervalStart = baseDate.AddHours(i);
                DateTime intervalEnd = intervalStart.AddHours(1);
                var data = sourceData.FirstOrDefault(x => x.ReportedTime >= intervalStart && x.ReportedTime < intervalEnd);//匹配时间
                if ( data != null)
                {
                    sortedList[i] = data;
                }
            }
            return sortedList;
        }
        //根据时间排序数据-操作员输入数据
        private static List<OperatorInputData> SortDataByTime(List<OperatorInputData> OperatorInputData, DateTime baseDate, int maxCount)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<OperatorInputData>(new OperatorInputData[maxCount]);
            for (int i = 0; i < maxCount; i++)
            {
                DateTime intervalStart = baseDate.AddHours(i);
                DateTime intervalEnd = intervalStart.AddHours(1);
                var data = OperatorInputData.FirstOrDefault(x => x.ReportedTime >= intervalStart && x.ReportedTime < intervalEnd);//匹配时间
                if (data != null)
                {
                    sortedList[i] = data;
                }
            }
            return sortedList;
        }

    }
}
