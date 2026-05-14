using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using static FastExpressionCompiler.ExpressionCompiler;

namespace CenterBackend.Services
{
    public class DashboardService(IReportRepository<SourceData> sourceData, IReportRepository<OperatorInputData> operatorInputData) : IDashboardService
    {
        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;

        // 获取查询时间范围：昨日8:00到今日8:00
        private static (DateTime start, DateTime end) GetQueryTimeRange(DateTime time)
        {
            var startTime = time.AddDays(-1).Date.AddHours(8);
            var endTime = startTime.AddHours(24);
            return (startTime, endTime);
        }
        private static float? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())// 空数据校验
                return null;
            var validValues = data.Select(selector).OfType<float>();
            float sum = 0f;
            int count = 0;
            foreach (var value in validValues)
            {
                sum += value;
                count++;
            }
            return count > 0 ? sum / count : (float?)null;
        }
        private static float? CalculateFirstLastDifference<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())
                return null;

            var nonNullValues = data
                .Select(selector)
                .Where(x => x.HasValue)
                .Select(x => x.GetValueOrDefault())
                .ToList();
            if (nonNullValues.Count < 2)
                return null;
            float firstValue = nonNullValues.First();//计算差值
            float lastValue = nonNullValues.Last();
            float difference = lastValue - firstValue;
            return difference;
        }
        private static float? CalculateSum<T>(IEnumerable<T> data, Func<T, float?> selector)//非null值的总和
        {
            if (data == null || !data.Any())//空数据校验
                return null;
            var nonNullValues = data        //筛选非null的float值
                .Select(selector)           // 提取float?字段
                .Where(x => x.HasValue)     // 过滤掉null值
                .Select(x => x.GetValueOrDefault())      // 转换为float(非可空)
                .ToList();
            return nonNullValues.Count != 0 ? nonNullValues.Sum() : (float?)null;//计算总和
        }
        //从SourceData中查询数据生成一条曲线dto
        private async Task<LineChartDataDto> GetLineSourceDataAsync(DateTime currentTime, Func<SourceData, float?> valueSelector)
        {
            var (startTime, endTime) = GetQueryTimeRange(currentTime);
            endTime = currentTime;

            int totalHours = (int)Math.Ceiling((endTime - startTime).TotalHours);
            string[] xAxis = Enumerable.Range(0, totalHours)
                                       .Select(i => ((i + 8) % 24).ToString())
                                       .ToArray();

            List<SourceData> dataList = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
            if (dataList == null || dataList.Count == 0)
                return new LineChartDataDto();

            LineChartDataDto lineChartDataDto = new()
            {
                XAxis = xAxis,
                Series = new List<LineChartSeriesDto>
                {
                    new LineChartSeriesDto()
                    {
                        Name = "Series1",
                        Data = new float?[totalHours + 1]
                    }
                }
            }; 
            float?[] seriesData = lineChartDataDto.Series[0].Data;

            foreach (var dataItem in dataList)
            {
                int hourDiff = (int)Math.Floor((dataItem.ReportedTime - startTime).TotalHours);
                if (hourDiff < 0 || hourDiff >= totalHours)
                    continue;
                float? value = valueSelector(dataItem);
                if (value.HasValue)
                {
                    seriesData[hourDiff] = value;
                }
            }
            return lineChartDataDto;
        }

        //从OperatorInputData中查询数据生成一条曲线dto
        private async Task<LineChartDataDto> GetLineOperateDataAsync(DateTime currentTime, Func<OperatorInputData, float?> valueSelector)
        {
            var (startTime, endTime) = GetQueryTimeRange(currentTime);
            endTime = currentTime;

            int totalHours = (int)Math.Ceiling((endTime - startTime).TotalHours);
            string[] xAxis = Enumerable.Range(0, totalHours)
                                       .Select(i => (i % 24).ToString())
                                       .ToArray();

            List<OperatorInputData> dataList = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);
            if (dataList == null || dataList.Count == 0)
                return new LineChartDataDto();

            LineChartDataDto lineChartDataDto = new()
            {
                XAxis = xAxis,
                Series = new List<LineChartSeriesDto>
                {
                    new LineChartSeriesDto()
                    {
                        Name = "Series1",
                        Data = new float?[totalHours + 1]
                    }
                }
            };
            float?[] seriesData = lineChartDataDto.Series[0].Data;
            foreach (var dataItem in dataList)
            {
                int hourDiff = (int)Math.Floor((dataItem.ReportedTime - startTime).TotalHours);
                if (hourDiff < 0 || hourDiff >= totalHours)
                    continue;
                float? value = valueSelector(dataItem);
                if (value.HasValue)
                {
                    seriesData[hourDiff] = value;
                }
            }
            return lineChartDataDto;
        }

        public async Task<CoreChartDto> GetPage1CoreChart1()
        {
            (var startTime, var endtime) = GetQueryTimeRange(DateTime.Now);//记录上一个班次的数据
            List<SourceData> dataList = await _sourceData.GetByDateTimeRangeAsync(startTime, endtime);
            CoreChartDto coreChartDto = new();
            if (dataList == null || dataList.Count == 0) 
                return coreChartDto;

            coreChartDto.Card1 = CalculateAverage(dataList, x => x.Cell19);//羟基流量
            coreChartDto.Card2 = CalculateAverage(dataList, x => x.Cell13);//气氨流量
            coreChartDto.Card3 = CalculateAverage(dataList, x => x.Cell23);//摩尔比
            coreChartDto.Card4 = CalculateAverage(dataList, x => x.Cell15);//配料蒸汽
            coreChartDto.Card5 = CalculateAverage(dataList, x => x.Cell26);//热点温度

            return coreChartDto;
        }

        public async Task<LineChartDataDto> GetPage1LineChart1()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell19);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "羟基流量(L/h)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage1LineChart2()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell13);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "气氨流量(kg/h)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage1LineChart3()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell23);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "摩尔比";
            return result;
        }

        public async Task<LineChartDataDto> GetPage1LineChart4()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell15);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "配料蒸汽流量(m³/h)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage1LineChart5()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell26);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "反应器热点温度(℃)";
            return result;
        }

        public async Task<CoreChartDto> GetPage2CoreChart1()
        {
            (var startTime, var endtime) = GetQueryTimeRange(DateTime.Now);
            List<SourceData> dataList = await _sourceData.GetByDateTimeRangeAsync(startTime, endtime);
            List<OperatorInputData> dataList2 = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endtime);
            CoreChartDto coreChartDto = new();
            if (dataList == null || dataList.Count == 0)
            {
                return coreChartDto;
            }
            else
            {
                coreChartDto.Card1 = CalculateAverage(dataList, x => x.Cell66);//一次结晶温度

                coreChartDto.Card3 = CalculateFirstLastDifference(dataList, x => x.Cell145);//降膜蒸发进料量FT-101
                coreChartDto.Card4 = CalculateFirstLastDifference(dataList, x => x.Cell146);//降膜蒸发出料量FR-102
                coreChartDto.Card5 = CalculateAverage(dataList, x => x.Cell144);//降膜蒸发温度
                coreChartDto.Card6 = CalculateAverage(dataList, x => x.Cell122);//二次结晶温度
            }
            if (dataList2 == null || dataList2.Count == 0)
            {
                return coreChartDto;
            }
            else
            {
                coreChartDto.Card2 = CalculateSum(dataList2, x => x.Cell26);//一次结晶产量
                coreChartDto.Card7 = CalculateSum(dataList2, x => x.Cell36);//二次结晶温度
            }
            return coreChartDto;
        }

        public async Task<LineChartDataDto> GetPage2LineChart1()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell66);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "一次结晶温度(℃)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage2LineChart2()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell144);//新增的数据采集点位，降膜蒸发温度
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "降膜蒸发温度(℃)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage2LineChart3()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell122);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "二次结晶温度(℃)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage2LineChart4()
        {
            var result = await GetLineOperateDataAsync(DateTime.Now, x => x.Cell26);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "一次结晶产量(kg)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage2LineChart5()
        {
            var result = await GetLineOperateDataAsync(DateTime.Now, x => x.Cell36);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "二次结晶产量(kg)";
            return result;
        }

        public async Task<CoreChartDto> GetPage3CoreChart1()
        {
            (var startTime, var endtime) = GetQueryTimeRange(DateTime.Now);
            List<SourceData> dataList = await _sourceData.GetByDateTimeRangeAsync(startTime, endtime);
            List<OperatorInputData> dataList2 = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endtime);
            CoreChartDto coreChartDto = new();
            if (dataList == null || dataList.Count == 0)
            {
                return coreChartDto;
            }
            else
            {
                coreChartDto.Card4 = CalculateFirstLastDifference(dataList, x => x.Cell134);//废液排放
            }
            if (dataList2 == null || dataList2.Count == 0)
            {
                return coreChartDto;
            }
            else
            {
                float? difference;
                //低压蒸汽消耗
                difference = CalculateFirstLastDifference(dataList, x => x.Cell71);
                coreChartDto.Card1 = difference > 0 ? difference : null;
                //中压蒸汽消耗
                difference = CalculateFirstLastDifference(dataList, x => x.Cell72);
                coreChartDto.Card2 = difference > 0 ? difference : null;
                //电量消耗
                difference = CalculateFirstLastDifference(dataList, x => x.Cell73);
                coreChartDto.Card3 = difference > 0 ? difference : null;
            }

            return coreChartDto;
        }

        public async Task<LineChartDataDto> GetPage3LineChart1()
        {
            var result = await GetLineOperateDataAsync(DateTime.Now, x => x.Cell71);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "低压蒸汽消耗(t)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage3LineChart2()
        {
            var result = await GetLineOperateDataAsync(DateTime.Now, x => x.Cell72);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "中压蒸汽消耗(t)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage3LineChart3()
        {
            var result = await GetLineOperateDataAsync(DateTime.Now, x => x.Cell73);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "电量消耗(kwh)";
            return result;
        }

        public async Task<LineChartDataDto> GetPage3LineChart4()
        {
            var result = await GetLineSourceDataAsync(DateTime.Now, x => x.Cell133);
            if (result.Series != null && result.Series.Count > 0)
                result.Series[0].Name = "废液排放流量(m³)";
            return result;
        }

    }


}
