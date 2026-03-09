using CenterBackend.IServices;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using NPOI.HPSF;
using NPOI.OpenXmlFormats.Shared;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;

namespace CenterBackend.Services
{
    public class DataToViewService( IReportRepository<SourceData> sourceData,
                                    IReportRepository<OperatorInputData> operatorInputData,
                                    IReportRepository<CalculatedData> calculatedData) : IDataToViewService
    {

        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;
        private readonly IReportRepository<CalculatedData> _calculatedData = calculatedData;

        public async Task<bool> DayGetMapDataAsync(DayWorkBook DayWorkBook)
        {
            var startTime = DayWorkBook.ReportedTime.Date.AddHours(8);
            var endTime = DayWorkBook.ReportedTime.Date.AddDays(1).AddHours(8);

            var sourceData = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
            if (sourceData == null || sourceData.Count == 0)//未查到数据
                return false;
            var operatorInputData = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);

            DayWorkBook.DaySheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();
            DayWorkBook.NightSheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();

            var baseTime = startTime;
            var dataPart1 = SortDataByTime(sourceData, baseTime, 25);//原始数据
            var dataPart2 = SortDataByTime(operatorInputData, baseTime, 25);//人工录入数据

            List<SourceData> source1 = dataPart1.Take(13).ToList();
            List<OperatorInputData> source2 = dataPart2.Take(13).ToList();
            DayMoveData(DayWorkBook.DaySheet, source1, source2);//白班

            source1 = dataPart1.Skip(12).Take(13).ToList();
            source2 = dataPart2.Skip(12).Take(13).ToList();
            DayMoveData(DayWorkBook.NightSheet, source1, source2);//夜班

            return true;
        }

        public async Task<bool> MonthGetMapDataAsync(MonthWorkBook monthWorkBook)
        {
            return false;
        }

        public async Task<bool> YearGetMapDataAsync(YearWorkBook yearWorkBook)
        {
            return false;
        }

        public async Task<bool> WeekGetMapDataAsync(WeekWorkBook WeekWorkBook)
        {
            await WeekMoveDataSheet2Async(WeekWorkBook);
            await WeekMoveDataSheet3Async(WeekWorkBook);
            await WeekMoveDataSheet4Async(WeekWorkBook);
            await WeekMoveDataSheet5Async(WeekWorkBook);
            await WeekMoveDataSheet6Async(WeekWorkBook);
            await WeekMoveDataSheet7Async(WeekWorkBook);
            await WeekMoveDataSheet8Async(WeekWorkBook);
            await WeekMoveDataSheet9Async(WeekWorkBook);
            await WeekMoveDataSheet10Async(WeekWorkBook);
            await WeekMoveDataSheet11Async(WeekWorkBook);
            await WeekMoveDataSheet12Async(WeekWorkBook);
            await WeekMoveDataSheet13Async(WeekWorkBook);
            return true;
        }
        /***********************数据处理***********************/
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
                if (data != null)
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
        private static void DayMoveData(List<DayWorkSheet> DayWorkSheet, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            var target = DayWorkSheet;
            var source1 = sourceData;
            var source2 = operatorInputData;
            //原始数据的映射
            for (int i = 0; i < 13; i++)
            {
                if (source1 == null || source1[i] == null)
                    continue;
                target[i].Cell1 = source1[i].Cell1;
                target[i].Cell2 = source1[i].Cell2;
                target[i].Cell3 = source1[i].Cell3;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell4;
                    var prevData = source1[i - 1]?.Cell4;// 可能为 null
                    if (currentVal != null && prevData != null)
                        target[i].Cell4 = (currentVal - prevData) / 1000;
                }
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell5;
                    var prevData = source1[i - 1]?.Cell5;
                    if (currentVal != null && prevData != null)
                        target[i].Cell5 = (currentVal - prevData) / 1000;
                }
                target[i].Cell6 = source1[i].Cell6;
                target[i].Cell7 = source1[i].Cell7;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell8;
                    var prevData = source1[i - 1]?.Cell8;
                    if (currentVal != null && prevData != null)
                        target[i].Cell8 = (currentVal - prevData) / 1000;
                }
                target[i].Cell9 = source1[i].Cell9;
                target[i].Cell10 = source1[i].Cell10;
                target[i].Cell11 = source1[i].Cell11;
                target[i].Cell12 = source1[i].Cell12;
                target[i].Cell13 = source1[i].Cell13;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell14;
                    var prevData = source1[i - 1]?.Cell14;
                    if (currentVal != null && prevData != null)
                        target[i].Cell14 = (currentVal - prevData) / 1000;
                }
                target[i].Cell15 = source1[i].Cell15;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell16;
                    var prevData = source1[i - 1]?.Cell16;
                    if (currentVal != null && prevData != null)
                        target[i].Cell16 = (currentVal - prevData) / 1000;
                }
                target[i].Cell17 = source1[i].Cell17;
                target[i].Cell18 = source1[i].Cell18;
                target[i].Cell19 = source1[i].Cell19;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell20;
                    var prevData = source1[i - 1]?.Cell20;
                    if (currentVal != null && prevData != null)
                        target[i].Cell20 = (currentVal - prevData) / 1000;
                }
                target[i].Cell21 = source1[i].Cell21;

                target[i].Cell22 = source1[i].Cell22 < 2 ? source1[i].Cell22 : -1;//摩尔比 小于2 直接出 
                if (i == 12)// 最后一个值
                    target[i].Cell23 = source1[i].Cell23;
                target[i].Cell24 = source1[i].Cell24;
                target[i].Cell25 = source1[i].Cell25;
                target[i].Cell26 = source1[i].Cell26;
                target[i].Cell27 = source1[i].Cell27;
                target[i].Cell28 = source1[i].Cell28;
                //人工检测数据
                //target[i].Cell29 = source1[i].Cell29;
                //target[i].Cell30 = source1[i].Cell30;
                //target[i].Cell31 = source1[i].Cell31;
                //target[i].Cell32 = source1[i].Cell32;
                //target[i].Cell33 = source1[i].Cell33;
                //target[i].Cell34 = source1[i].Cell34;
                //target[i].Cell35 = source1[i].Cell35;

                target[i].Cell36 = source1[i].Cell36;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell37;
                    var prevData = source1[i - 1]?.Cell37;
                    if (currentVal != null && prevData != null)
                        target[i].Cell37 = (currentVal - prevData) / 1000;
                }
                target[i].Cell38 = source1[i].Cell38;
                target[i].Cell39 = source1[i].Cell39;
                target[i].Cell40 = source1[i].Cell40;
                target[i].Cell41 = source1[i].Cell41;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell42;
                    var prevData = source1[i - 1]?.Cell42;
                    if (currentVal != null && prevData != null)
                        target[i].Cell42 = (currentVal - prevData) / 1000;
                }
                //预留数据
                //target[i].Cell43 = source1[i].Cell43;
                //target[i].Cell44 = source1[i].Cell44;
                //target[i].Cell45 = source1[i].Cell45;
                //target[i].Cell46 = source1[i].Cell46;
                //target[i].Cell47 = source1[i].Cell47;
                //target[i].Cell48 = source1[i].Cell48;
                //target[i].Cell49 = source1[i].Cell49;
                //target[i].Cell50 = source1[i].Cell50;

                target[i].Cell51 = source1[i].Cell51;
                target[i].Cell52 = source1[i].Cell52;
                target[i].Cell53 = source1[i].Cell53;
                target[i].Cell54 = source1[i].Cell54;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell55;
                    var prevData = source1[i - 1]?.Cell55;
                    if (currentVal != null && prevData != null)
                        target[i].Cell55 = (currentVal - prevData) / 1000;
                }
                //人工检测数据
                //target[i].Cell56 = source1[i].Cell56;
                //target[i].Cell57 = source1[i].Cell57;
                //target[i].Cell58 = source1[i].Cell58;
                //target[i].Cell59 = source1[i].Cell59;
                //target[i].Cell60 = source1[i].Cell60;

                target[i].Cell61 = source1[i].Cell61;
                target[i].Cell62 = source1[i].Cell62;
                target[i].Cell63 = source1[i].Cell63;
                target[i].Cell64 = source1[i].Cell64;
                target[i].Cell65 = source1[i].Cell65;
                target[i].Cell66 = source1[i].Cell66;
                target[i].Cell67 = source1[i].Cell67;
                target[i].Cell68 = source1[i].Cell68;
                target[i].Cell69 = source1[i].Cell69;
                target[i].Cell70 = source1[i].Cell70;
                target[i].Cell71 = source1[i].Cell71;
                target[i].Cell72 = source1[i].Cell72;
                target[i].Cell73 = source1[i].Cell73;
                target[i].Cell74 = source1[i].Cell74;
                target[i].Cell75 = source1[i].Cell75;
                target[i].Cell76 = source1[i].Cell76;
                target[i].Cell77 = source1[i].Cell77;
                target[i].Cell78 = source1[i].Cell78;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell79;
                    var prevData = source1[i - 1]?.Cell79;
                    if (currentVal != null && prevData != null)
                        target[i].Cell79 = (currentVal - prevData) / 1000;
                }
                target[i].Cell80 = source1[i].Cell80;
                target[i].Cell81 = source1[i].Cell81;
                //人工检测数据
                //target[i].Cell82 = source1[i].Cell82;
                //target[i].Cell83 = source1[i].Cell83;
                //target[i].Cell84 = source1[i].Cell84;
                //target[i].Cell85 = source1[i].Cell85;
                //target[i].Cell86 = source1[i].Cell86;
                //target[i].Cell87 = source1[i].Cell87;

                target[i].Cell88 = source1[i].Cell88;
                target[i].Cell89 = source1[i].Cell89;
                target[i].Cell90 = source1[i].Cell90;
                target[i].Cell91 = source1[i].Cell91;
                target[i].Cell92 = source1[i].Cell92;
                //预留数据
                //target[i].Cell93 = source1[i].Cell93;
                //target[i].Cell94 = source1[i].Cell94;
                //target[i].Cell95 = source1[i].Cell95;
                //target[i].Cell96 = source1[i].Cell96;
                //target[i].Cell97 = source1[i].Cell97;
                //target[i].Cell98 = source1[i].Cell98;
                //target[i].Cell99 = source1[i].Cell99;
                //target[i].Cell100 = source1[i].Cell100;
                target[i].Cell101 = source1[i].Cell101;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell102;
                    var prevData = source1[i - 1]?.Cell102;
                    if (currentVal != null && prevData != null)
                        target[i].Cell102 = (currentVal - prevData) / 1000;
                }
                target[i].Cell103 = source1[i].Cell103;
                target[i].Cell104 = source1[i].Cell104;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell105;
                    var prevData = source1[i - 1]?.Cell105;
                    if (currentVal != null && prevData != null)
                        target[i].Cell105 = (currentVal - prevData) / 1000;
                }
                target[i].Cell106 = source1[i].Cell106;
                target[i].Cell107 = source1[i].Cell107;
                target[i].Cell108 = source1[i].Cell108;
                target[i].Cell109 = source1[i].Cell109;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell110;
                    var prevData = source1[i - 1]?.Cell110;
                    if (currentVal != null && prevData != null)
                        target[i].Cell110 = (currentVal - prevData) / 1000;
                }
                target[i].Cell111 = source1[i].Cell111;
                target[i].Cell112 = source1[i].Cell112;
                target[i].Cell113 = source1[i].Cell113;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell114;
                    var prevData = source1[i - 1]?.Cell114;
                    if (currentVal != null && prevData != null)
                        target[i].Cell114 = (currentVal - prevData) / 1000;
                }
                target[i].Cell115 = source1[i].Cell115;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell116;
                    var prevData = source1[i - 1]?.Cell116;
                    if (currentVal != null && prevData != null)
                        target[i].Cell116 = (currentVal - prevData) / 1000;
                }
                target[i].Cell117 = source1[i].Cell117;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell118;
                    var prevData = source1[i - 1]?.Cell118;
                    if (currentVal != null && prevData != null)
                        target[i].Cell118 = (currentVal - prevData) / 1000;
                }
                target[i].Cell119 = source1[i].Cell119;
                target[i].Cell120 = source1[i].Cell120;
                target[i].Cell121 = source1[i].Cell121;
                target[i].Cell122 = source1[i].Cell122;
                target[i].Cell123 = source1[i].Cell123;
                target[i].Cell124 = source1[i].Cell124;
                target[i].Cell125 = source1[i].Cell125;
                target[i].Cell126 = source1[i].Cell126;
                target[i].Cell127 = source1[i].Cell127;
                target[i].Cell128 = source1[i].Cell128;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell129;
                    var prevData = source1[i - 1]?.Cell129;
                    if (currentVal != null && prevData != null)
                        target[i].Cell129 = (currentVal - prevData) / 1000;
                }
                target[i].Cell130 = source1[i].Cell130;
                target[i].Cell131 = source1[i].Cell131;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell132;
                    var prevData = source1[i - 1]?.Cell132;
                    if (currentVal != null && prevData != null)
                        target[i].Cell132 = (currentVal - prevData) / 1000;
                }
                target[i].Cell133 = source1[i].Cell133;
                target[i].Cell134 = source1[i].Cell134;
                //人工检测数据
                //target[i].Cell135 = source1[i].Cell135;
                //target[i].Cell136 = source1[i].Cell136;
                //target[i].Cell137 = source1[i].Cell137;
                //target[i].Cell138 = source1[i].Cell138;
                //target[i].Cell139 = source1[i].Cell139;
                //target[i].Cell140 = source1[i].Cell140;
                //target[i].Cell141 = source1[i].Cell141;
                //预留数据
                //target[i].Cell142 = source1[i].Cell142;
                //target[i].Cell143 = source1[i].Cell143;
                //target[i].Cell144 = source1[i].Cell144;
                //target[i].Cell145 = source1[i].Cell145;
                //target[i].Cell146 = source1[i].Cell146;
                //target[i].Cell147 = source1[i].Cell147;
                //target[i].Cell148 = source1[i].Cell148;
                //target[i].Cell149 = source1[i].Cell149;
                //target[i].Cell150 = source1[i].Cell150;
            }
            //人工输入数据的映射
            for (int i = 0; i < 13; i++) {
                if (source2 == null || source2[i] == null)
                    continue;
                //人工检测数据
                target[i].Cell29 = source2[i].Cell1;
                target[i].Cell30 = source2[i].Cell2;
                target[i].Cell31 = source2[i].Cell3;
                target[i].Cell32 = source2[i].Cell4;
                target[i].Cell33 = source2[i].Cell5;
                target[i].Cell34 = source2[i].Cell6;
                target[i].Cell35 = source2[i].Cell7;
                //人工检测数据
                target[i].Cell56 = source2[i].Cell11;
                target[i].Cell57 = source2[i].Cell12;
                target[i].Cell58 = source2[i].Cell13;
                target[i].Cell59 = source2[i].Cell14;
                target[i].Cell60 = source2[i].Cell15;
                //人工检测数据
                target[i].Cell82 = source2[i].Cell21;
                target[i].Cell83 = source2[i].Cell22;
                target[i].Cell84 = source2[i].Cell23;
                target[i].Cell85 = source2[i].Cell24;
                target[i].Cell86 = source2[i].Cell25;
                target[i].Cell87 = source2[i].Cell26;
                //人工检测数据
                target[i].Cell135 = source2[i].Cell31;
                target[i].Cell136 = source2[i].Cell32;
                target[i].Cell137 = source2[i].Cell33;
                target[i].Cell138 = source2[i].Cell34;
                target[i].Cell139 = source2[i].Cell35;
                target[i].Cell140 = source2[i].Cell36;
                target[i].Cell141 = source2[i].Cell37;
            }
        }

        private static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday; 
            if (diff < 0) diff += 7;
            return dt.AddDays(-diff).Date;
        }

        private static float? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())// 空数据校验
                return null;
            var nonNullValues = data// 筛选非null的float值并计算平均值
                .Select(selector)
                .Where(x => x.HasValue)
                .Select(x => x.GetValueOrDefault())
                .ToList();
            return nonNullValues.Count != 0 ? nonNullValues.Average() : (float?)null;// 计算平均值
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

        /// <param name="mode">mode：模式开关（true = 区间校验 [min,max]，false = 上限校验 [≤max]）</param>
        private static float? CalculateQualifiedRate<T>(IEnumerable<T> data, Func<T, float?> selector,bool mode, float  qualifiedValue, float qualifiedValuediff)
        {
            var maxQualifiedValue = qualifiedValue + qualifiedValuediff;
            var minQualifiedValue = qualifiedValue - qualifiedValuediff;
            if (data == null || !data.Any())// 空数据校验
                return null;
            var nonNullValues = data// 筛选非null的float值并计算平均值
                .Select(selector)
                .Where(x => x.HasValue)
                .Select(x => x.GetValueOrDefault())
                .ToList();
            if (nonNullValues.Count == 0)
                return null;

            int qualifiedCount = mode
                // mode=true：值在 [min, max] 区间内为合格
                ? nonNullValues.Count(value => value >= minQualifiedValue && value <= maxQualifiedValue)
                // mode=false：值 ≤ max 为合格
                : nonNullValues.Count(value => value <= maxQualifiedValue);
            float qualifiedRate = (qualifiedCount / (float)nonNullValues.Count) * 100f;// 计算合格利率（合格数/总有效数 * 100%），保留3位小数
            return (float)Math.Round(qualifiedRate, 3);
        }
        private async Task<bool> WeekMoveDataSheet2Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet2 = Enumerable.Range(1, 3).Select(_ => new WorkSheet2()).ToList();

            DateTime startTime ;
            DateTime endTime ;
            List<SourceData> sourceData = [];
            DateTime currentMonthFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = currentMonthFirstDay.AddDays(-14);
                    endTime = startTime.AddDays(7);
                }
                else if(i == 1)
                {
                    startTime = currentMonthFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentMonthFirstDay;
                    endTime = startTime.AddDays(7);
                }

                sourceData = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
                if (sourceData.Count == 0)//无数据则跳过
                    continue;
                if (sourceData != null && sourceData.Count != 0)
                {
                    WeekWorkBook.WorkSheet2[i].Cell1 = CalculateAverage(sourceData, x => x.Cell13);
                    WeekWorkBook.WorkSheet2[i].Cell2 = CalculateAverage(sourceData, x => x.Cell15);
                    WeekWorkBook.WorkSheet2[i].Cell3 = CalculateAverage(sourceData, x => x.Cell19);
                    WeekWorkBook.WorkSheet2[i].Cell4 = CalculateAverage(sourceData, x => x.Cell22);
                    WeekWorkBook.WorkSheet2[i].Cell5 = CalculateAverage(sourceData, x => x.Cell24);
                    WeekWorkBook.WorkSheet2[i].Cell6 = CalculateAverage(sourceData, x => x.Cell25);
                    WeekWorkBook.WorkSheet2[i].Cell7 = CalculateAverage(sourceData, x => x.Cell26);
                    WeekWorkBook.WorkSheet2[i].Cell8 = CalculateAverage(sourceData, x => x.Cell27);
                    WeekWorkBook.WorkSheet2[i].Cell9 = CalculateAverage(sourceData, x => x.Cell28);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet3Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet3 = Enumerable.Range(1, 3).Select(_ => new WorkSheet3()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Year, WeekWorkBook.ReportedTime.Month, 1);//本月第一天
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 1 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet3[i].Cell1 = CalculateAverage(calculatedData, x => x.Cell151);
                    WeekWorkBook.WorkSheet3[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell152);
                    WeekWorkBook.WorkSheet3[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell153);
                    WeekWorkBook.WorkSheet3[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell154);
                    WeekWorkBook.WorkSheet3[i].Cell5 = CalculateAverage(calculatedData, x => x.Cell156);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet4Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet4 = Enumerable.Range(1, 3).Select(_ => new WorkSheet4()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-14);
                    endTime = startTime.AddDays(1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 2 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet4[i].Cell1 = CalculateAverage(calculatedData, x => x.Cell161);
                    WeekWorkBook.WorkSheet4[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell162);
                    WeekWorkBook.WorkSheet4[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell163);
                    WeekWorkBook.WorkSheet4[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell164);
                    WeekWorkBook.WorkSheet4[i].Cell5 = CalculateAverage(calculatedData, x => x.Cell165);

                    WeekWorkBook.WorkSheet4[i].Cell6 = CalculateAverage(calculatedData, x => x.Cell92);
                    WeekWorkBook.WorkSheet4[i].Cell7 = CalculateAverage(calculatedData, x => x.Cell106);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet5Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet5 = Enumerable.Range(1, 3).Select(_ => new WorkSheet5()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Year, 1, 1);//本年第一天
                    endTime = startTime.AddYears(1).AddDays(-1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 1 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet5[i].Cell1 = CalculateAverage(calculatedData, x => x.Cell121);
                    WeekWorkBook.WorkSheet5[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell122);
                    WeekWorkBook.WorkSheet5[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell127);
                    WeekWorkBook.WorkSheet5[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell128);

                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet6Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet6 = Enumerable.Range(1, 3).Select(_ => new WorkSheet6()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-14);
                    endTime = startTime.AddDays(1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 2 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet6[i].Cell1 = CalculateAverage(calculatedData, x => x.Cell83);
                    WeekWorkBook.WorkSheet6[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell84);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet7Async(WeekWorkBook WeekWorkBook)
        {
            return true;
        }
        private async Task<bool> WeekMoveDataSheet8Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet8 = Enumerable.Range(1, 9).Select(_ => new WorkSheet8()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 7; i++)
            {
                startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(i);
                endTime = startTime.AddDays(1);
                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet8[i].TimePoint = startTime;
                    WeekWorkBook.WorkSheet8[i].Cell1 = calculatedData[0].Cell161;
                    WeekWorkBook.WorkSheet8[i].Cell2 = calculatedData[0].Cell164;
                    WeekWorkBook.WorkSheet8[i].Cell3 = calculatedData[0].Cell167;
                    WeekWorkBook.WorkSheet8[i].Cell4 = calculatedData[0].Cell191;
                    WeekWorkBook.WorkSheet8[i].Cell5 = calculatedData[0].Cell193;
                    WeekWorkBook.WorkSheet8[i].Cell6 = calculatedData[0].Cell194;
                    WeekWorkBook.WorkSheet8[i].Cell7 = calculatedData[0].Cell195;
                    WeekWorkBook.WorkSheet8[i].Cell8 = calculatedData[0].Cell161;//转化率？？
                    WeekWorkBook.WorkSheet8[i].Cell9 = calculatedData[0].Cell161;//收率？？
                    WeekWorkBook.WorkSheet8[i].Cell10 = calculatedData[0].Cell197;
                    WeekWorkBook.WorkSheet8[i].Cell11 = calculatedData[0].Cell197;//废液二睛含量？？
                }
            }
            for (var i = 7; i < 9; i++)
            {
                if (i == 7)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }
                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 8 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet8[i].Cell1 = CalculateAverage(calculatedData, x => x.Cell13);
                    WeekWorkBook.WorkSheet8[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell15);
                    WeekWorkBook.WorkSheet8[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell19);
                    WeekWorkBook.WorkSheet8[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell22);
                    WeekWorkBook.WorkSheet8[i].Cell5 = CalculateAverage(calculatedData, x => x.Cell24);
                    WeekWorkBook.WorkSheet8[i].Cell6 = CalculateAverage(calculatedData, x => x.Cell25);
                    WeekWorkBook.WorkSheet8[i].Cell7 = CalculateAverage(calculatedData, x => x.Cell26);
                    WeekWorkBook.WorkSheet8[i].Cell8 = CalculateAverage(calculatedData, x => x.Cell27);
                    WeekWorkBook.WorkSheet8[i].Cell9 = CalculateAverage(calculatedData, x => x.Cell28);
                    WeekWorkBook.WorkSheet8[i].Cell10 = CalculateAverage(calculatedData, x => x.Cell28);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet9Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet9 = Enumerable.Range(1, 2).Select(_ => new WorkSheet9()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 2; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 1 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet9[i].Cell1 = CalculateFirstLastDifference(calculatedData, x => x.Cell102);//差值
                    WeekWorkBook.WorkSheet9[i].Cell2 = CalculateFirstLastDifference(calculatedData, x => x.Cell114);
                    WeekWorkBook.WorkSheet9[i].Cell3 = CalculateFirstLastDifference(calculatedData, x => x.Cell112);
                    WeekWorkBook.WorkSheet9[i].Cell4 = CalculateFirstLastDifference(calculatedData, x => x.Cell110);
                    //WeekWorkBook.WorkSheet9[i].Cell5 = CalculateAverage(calculatedData, x => x.Cell211);//低温蒸发没有检测数据
                    //WeekWorkBook.WorkSheet9[i].Cell6 = CalculateAverage(calculatedData, x => x.Cell213);
                    //WeekWorkBook.WorkSheet9[i].Cell7 = CalculateAverage(calculatedData, x => x.Cell215);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet10Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet10 = Enumerable.Range(1, 2).Select(_ => new WorkSheet10()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 2; i++)
            {
                if (i == 0)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 1 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet10[i].Cell1 = CalculateFirstLastDifference(calculatedData, x => x.Cell104);//差值
                    WeekWorkBook.WorkSheet10[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell226);
                    WeekWorkBook.WorkSheet10[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell221);
                    WeekWorkBook.WorkSheet10[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell223);
                    WeekWorkBook.WorkSheet10[i].Cell5 = CalculateAverage(calculatedData, x => x.Cell225);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet11Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet11 = Enumerable.Range(1, 3).Select(_ => new WorkSheet11()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Date.Year, WeekWorkBook.ReportedTime.Date.Month, 1);
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime, 1);
                if (i == 2 && calculatedData.Count == 0)//本周无数据则退出
                    return false;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    WeekWorkBook.WorkSheet11[i].Cell1 = CalculateFirstLastDifference(calculatedData, x => x.Cell132);//差值
                    WeekWorkBook.WorkSheet11[i].Cell2 = CalculateAverage(calculatedData, x => x.Cell211);
                    WeekWorkBook.WorkSheet11[i].Cell3 = CalculateAverage(calculatedData, x => x.Cell213);
                    WeekWorkBook.WorkSheet11[i].Cell4 = CalculateAverage(calculatedData, x => x.Cell215);
                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet12Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet12 = Enumerable.Range(1, 3).Select(_ => new WorkSheet12()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<CalculatedData> calculatedData = [];
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Date.Year, WeekWorkBook.ReportedTime.Date.Month, 1);
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }
                else if (i == 1)
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddDays(-7);
                    endTime = startTime.AddDays(1);
                }
                else
                {
                    startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date);
                    endTime = startTime.AddDays(1);
                }

                calculatedData = await _calculatedData.GetByDateTimeRangeAsync(startTime, endTime);
                float? temp1, temp2;
                if (calculatedData != null && calculatedData.Count != 0)
                {
                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell20);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell1 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell4);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell2 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell37);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell3 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell230);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell4 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell16) * 0.180218f / 1000;
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell5 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell110);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell6 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell55)+ CalculateFirstLastDifference(calculatedData, x => x.Cell114);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell7 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                    temp1 = CalculateFirstLastDifference(calculatedData, x => x.Cell130);
                    temp2 = CalculateFirstLastDifference(calculatedData, x => x.Cell197);
                    WeekWorkBook.WorkSheet12[i].Cell8 = temp2 != 0 ? temp1 * temp1 / temp2 : null;

                }
            }
            return true;
        }
        private async Task<bool> WeekMoveDataSheet13Async(WeekWorkBook WeekWorkBook)
        {
            WeekWorkBook.WorkSheet13 = Enumerable.Range(1, 14).Select(_ => new WorkSheet13()).ToList();

            DateTime startTime;
            DateTime endTime;

            List<SourceData> sourceData = [];
            List<OperatorInputData> operatorInputData = [];
            for (var i = 0; i < 14; i++)
            {
                startTime = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8 + i * 12);
                endTime = startTime.AddHours(12);

                sourceData = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
                operatorInputData = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);
                if (sourceData != null && sourceData.Count != 0)
                {
                    WeekWorkBook.WorkSheet13[i].Cell1 = CalculateQualifiedRate(sourceData, x => x.Cell20, true, 0.5150f, 0.05f);
                    WeekWorkBook.WorkSheet13[i].Cell2 = CalculateQualifiedRate(sourceData, x => x.Cell20, true, 410F, 5f);
                    WeekWorkBook.WorkSheet13[i].Cell3 = CalculateQualifiedRate(sourceData, x => x.Cell20, true, 168f, 2f);
                    WeekWorkBook.WorkSheet13[i].Cell4 = CalculateQualifiedRate(sourceData, x => x.Cell20, false, 20, 0);
                    //WeekWorkBook.WorkSheet13[i].Cell5 = CalculateQualifiedRate(sourceData, x => x.Cell20, 0.5150f, 0.05f);//
                    WeekWorkBook.WorkSheet13[i].Cell6 = CalculateFirstLastDifference(operatorInputData, x => x.Cell47);//手动录入的产量
                    WeekWorkBook.WorkSheet13[i].Cell7 = CalculateFirstLastDifference(operatorInputData, x => x.Cell47);//手动录入的产量
                    WeekWorkBook.WorkSheet13[i].Cell8 = CalculateFirstLastDifference(operatorInputData, x => x.Cell47);//手动录入的产量
                    WeekWorkBook.WorkSheet13[i].Cell9 = CalculateFirstLastDifference(operatorInputData, x => x.Cell47);//手动录入的产量
                }
            }
            return true;
        }


    }
}
