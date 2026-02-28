using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;

namespace CenterBackend.Services
{
    public class DashboardService : IDashboardService
    {
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IReportRepository<CalculatedData> _calculatedData;
        public DashboardService(IReportRepository<SourceData> _SourceData, IReportRepository<CalculatedData> _CalculatedData)
        {
            this._sourceData = _SourceData;
            this._calculatedData = _CalculatedData;
        }
        public async Task<LineChartDataDto> getLineChartOne(DateTime time)
        {
            //时间范围昨日0点到现在
            var queryStartTime = time.AddHours(-24).Date;
            var queryEndTime = time;
            var totalHours = (int)Math.Ceiling((queryEndTime - queryStartTime).TotalHours); // 总小时数（X轴长度）

            // 生成X轴刻度：0~totalHours-1（可按需替换为实际时间文本，如「昨日8点」）
            string[] xAxis = Enumerable.Range(0, totalHours)
                                       .Select(i => (i % 24).ToString())
                                       .ToArray();

            var chartDataDto = new LineChartDataDto//[全null基础DTO]
            {
                XAxis = xAxis,
                Series = new List<LineChartSeriesDto>
                {
                    new LineChartSeriesDto
                    {
                        Name = "羟基进料流量",
                        Data = Enumerable.Repeat((float?)null, totalHours).ToArray() // 全null数组
                    }
                }
            };

            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(queryStartTime, queryEndTime)//查询
                                                 .ConfigureAwait(false);
            if (sourceDatas == null || !sourceDatas.Any())// 无数据
            {
                return chartDataDto; // 直接返回全null基础DTO
            }

            // 有数据：按「时间→X轴索引」映射，填充对应位置的非null值
            float?[] DataLine1 = chartDataDto.Series[0].Data;
            foreach (var sourceData in sourceDatas)
            {
                var hourDiff = (int)Math.Floor((sourceData.ReportedTime - queryStartTime).TotalHours); // 计算当前数据时间与查询开始时间的小时差 → 对应X轴索引
                if (hourDiff < 0 || hourDiff >= totalHours)// 索引越界校验：过滤异常时间数据，避免数组报错
                {
                    continue;
                }
                if (sourceData.Cell19 != null && sourceData.Cell19 is float TempValue1)
                {
                    DataLine1[hourDiff] = TempValue1 * 1000;
                }
            }
            return chartDataDto;
        }


        public async Task<LineChartDataDto> getLineCharTwo(DateTime time)
        {
            //时间范围昨日0点到现在
            var queryStartTime = time.AddHours(-24).Date;
            var queryEndTime = time;
            var totalHours = (int)Math.Ceiling((queryEndTime - queryStartTime).TotalHours); // 总小时数（X轴长度）

            // 生成X轴刻度：0~totalHours-1（可按需替换为实际时间文本，如「昨日8点」）
            string[] xAxis = Enumerable.Range(0, totalHours)
                                       .Select(i => (i % 24).ToString())
                                       .ToArray();

            var chartDataDto = new LineChartDataDto//[全null基础DTO]
            {
                XAxis = xAxis,
                Series = new List<LineChartSeriesDto>
                {
                    new LineChartSeriesDto
                    {
                        Name = "摩尔比",
                        Data = Enumerable.Repeat((float?)null, totalHours).ToArray() // 全null数组
                    },
                    new LineChartSeriesDto
                    {
                        Name = "累计配比",
                        Data = Enumerable.Repeat((float?)null, totalHours).ToArray() // 全null数组
                    }
                }
            };

            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(queryStartTime, queryEndTime)//查询
                                                 .ConfigureAwait(false);
            if (sourceDatas == null || !sourceDatas.Any())// 无数据
            {
                return chartDataDto; // 直接返回全null基础DTO
            }

            // 有数据：按「时间→X轴索引」映射，填充对应位置的非null值
            float?[] DataLine1 = chartDataDto.Series[0].Data;
            float?[] DataLine2 = chartDataDto.Series[1].Data;
            foreach (var sourceData in sourceDatas)
            {
                var hourDiff = (int)Math.Floor((sourceData.ReportedTime - queryStartTime).TotalHours); // 计算当前数据时间与查询开始时间的小时差 → 对应X轴索引
                if (hourDiff < 0 || hourDiff >= totalHours)// 索引越界校验：过滤异常时间数据，避免数组报错
                {
                    continue;
                }

                if (sourceData.Cell22 != null && sourceData.Cell22 is float TempValue1)
                {
                    DataLine1[hourDiff] = TempValue1;
                }

                if (sourceData.Cell23 != null && sourceData.Cell23 is float TempValue2)
                {
                    DataLine2[hourDiff] = TempValue2;
                }
            }
            return chartDataDto;
        }

        public async Task<LineChartDataDto> getLineCharThree(DateTime time)
        {
            //时间范围昨日0点到现在
            var queryStartTime = time.AddHours(-24).Date;
            var queryEndTime = time;
            var totalHours = (int)Math.Ceiling((queryEndTime - queryStartTime).TotalHours); // 总小时数（X轴长度）

            // 生成X轴刻度：0~totalHours-1（可按需替换为实际时间文本，如「昨日8点」）
            string[] xAxis = Enumerable.Range(0, totalHours)
                                       .Select(i => (i % 24).ToString())
                                       .ToArray();

            //[全null基础DTO]
            var chartDataDto = new LineChartDataDto
            {
                XAxis = xAxis,
                Series = new List<LineChartSeriesDto>
                {
                    new LineChartSeriesDto
                    {
                        Name = "羟基原料浓度1",
                        Data = Enumerable.Repeat((float?)null, totalHours).ToArray() // 全null数组
                    },
                    new LineChartSeriesDto
                    {
                        Name = "羟基配后浓度2",
                        Data = Enumerable.Repeat((float?)null, totalHours).ToArray() // 全null数组
                    }
                }
            };

            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(queryStartTime, queryEndTime)//查询
                                                 .ConfigureAwait(false);
            if (sourceDatas == null || !sourceDatas.Any())// 无数据
            {
                return chartDataDto; // 直接返回全null基础DTO
            }

            // 有数据：按「时间→X轴索引」映射，填充对应位置的非null值
            float?[] DataLine1 = chartDataDto.Series[0].Data;
            float?[] DataLine2 = chartDataDto.Series[1].Data;
            foreach (var sourceData in sourceDatas)
            {
                var hourDiff = (int)Math.Floor((sourceData.ReportedTime - queryStartTime).TotalHours); // 计算当前数据时间与查询开始时间的小时差 → 对应X轴索引
                if (hourDiff < 0 || hourDiff >= totalHours)// 索引越界校验：过滤异常时间数据，避免数组报错
                {
                    continue;
                }
                if (sourceData.Cell3 != null && sourceData.Cell3 is float TempValue1)
                {
                    DataLine1[hourDiff] = TempValue1;
                }

                if (sourceData.Cell6 != null && sourceData.Cell6 is float TempValue2)
                {
                    DataLine2[hourDiff] = TempValue2;
                }
            }
            return chartDataDto;
        }

        public async Task<List<PieChartItemDto>> getPieChart(DateTime time)
        {
            await Task.Delay(1);
            var pieChartItems = new List<PieChartItemDto>
        {
            new PieChartItemDto { Name = "类别A", Value = 40 },
            new PieChartItemDto { Name = "类别B", Value = 30 },
            new PieChartItemDto { Name = "类别C", Value = 20 },
            new PieChartItemDto { Name = "类别D", Value = 10 }
        };

            return pieChartItems;
        }

        public async Task<CoreChartDto> getCoreChart(DateTime time)
        {
            //var StartTime = time.AddDays(-1).Date;
            var StartTime = time.Date;
            StartTime = StartTime.AddHours(8).AddMinutes(30);
            var EndTime = time.Date;
            EndTime = EndTime.AddHours(8).AddMinutes(40);

            List<CalculatedData> dataList = await _calculatedData.GetByDateTimeRangeAsync(StartTime, EndTime, 1);

            var coreChartDto = new CoreChartDto
            {
                Yesterday = 0,
                Week = 0,
                Month = 0,
                Year = 0
            };
            if (dataList == null || !dataList.Any()) return coreChartDto;

            if (dataList[0].Cell3 is float value1) coreChartDto.Yesterday = value1;
            if (dataList[0].Cell6 is float value2) coreChartDto.Week = value2;
            if (dataList[0].Cell22 is float value3 && value3 < 2) coreChartDto.Month = value3; // 大于2 则表示值错误
            if (dataList[0].Cell23 is float value4) coreChartDto.Year = value4;
            return coreChartDto;

        }
    }


}
