using CenterBackend.IServices;
using CenterBackend.Models.CalculateData;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterUser.Repository;
using System.Linq;
using Microsoft.EntityFrameworkCore;

namespace CenterBackend.Services
{

    public class CalculatedAndSaveService(
                IReportRepository<SourceData> sourceData,
                IReportRepository<OperatorInputData> operatorInputData,
                IReportRecordRepository<ReportRecord> reportRecord,
                IReportRepository<CalculatedData> calculatedDatas,
                IReportUnitOfWork reportUnitOfWork) : ICalculatedAndSaveService
    {        
        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;
        private readonly IReportRecordRepository<ReportRecord> _reportRecord = reportRecord;
        private readonly IReportRepository<CalculatedData> _calculatedDatas = calculatedDatas;
        private readonly IReportUnitOfWork _reportUnitOfWork = reportUnitOfWork;

        public async Task<bool> DataAnalyses(ReportInfo ReportInfo)
        {
            switch (ReportInfo.SheetType)
            {
                case SheetType.DayReport:
                    return await DayDataCalculateAsync(ReportInfo);
                case SheetType.MonthReport:
                    break;
                case SheetType.YearReport:
                    break;
                case SheetType.WeekReport:
                    break;
                default:
                    break;
            }
            return false;
        }

        private async Task<bool> DayDataCalculateAsync( ReportInfo ReportInfo)
        {

            if (ReportInfo == null || ReportInfo.TimeStart == default)
            {
                return false;
            }
            var target = await _calculatedDatas.Db.AsQueryable()
                .Where(r => r.ReportedTime.Date == ReportInfo.TimeStart.Date && r.Type == 1)
                .FirstOrDefaultAsync();
            bool existRecord = false;
            if (target != null)//更新记录
            {
                existRecord = true;
                target.ReportedTime = ReportInfo.TimeStart.Date;
                target.LastChange = DateTime.Now;
                target.Type = 1;
            }
            else//插入记录
            {
                existRecord = false;
                target = new CalculatedData()
                {
                    ReportedTime = ReportInfo.TimeStart.Date,
                    LastChange = DateTime.Now,
                    Type = 1,
                };
            }

            List<SourceData> rawDataPart1 = await _sourceData.GetByDateTimeRangeAsync(ReportInfo.TimeStart, ReportInfo.TimeEnd);
            if (rawDataPart1 == null || rawDataPart1.Count == 0)
                return false;
            List<SourceData> dataListPart1 = SortDataByTime(rawDataPart1, ReportInfo.TimeStart);
            {
                // 平均值（使用通用方法）
                target.Cell1 = AverageOf(dataListPart1, x => x?.Cell1);
                target.Cell2 = AverageOf(dataListPart1, x => x?.Cell2);
                target.Cell3 = AverageOf(dataListPart1, x => x?.Cell3);

                // 差值（首/尾有值时计算）
                target.Cell4 = DifferenceOf(dataListPart1, x => x?.Cell4);
                target.Cell5 = DifferenceOf(dataListPart1, x => x?.Cell5);

                // 平均值
                target.Cell6 = AverageOf(dataListPart1, x => x?.Cell6);
                target.Cell7 = AverageOf(dataListPart1, x => x?.Cell7);

                // 差值
                target.Cell8 = DifferenceOf(dataListPart1, x => x?.Cell8);

                // 平均值
                target.Cell9 = AverageOf(dataListPart1, x => x?.Cell9);
                target.Cell10 = AverageOf(dataListPart1, x => x?.Cell10);
                target.Cell11 = AverageOf(dataListPart1, x => x?.Cell11);
                target.Cell12 = AverageOf(dataListPart1, x => x?.Cell12);
                target.Cell13 = AverageOf(dataListPart1, x => x?.Cell13);

                // 差值
                target.Cell14 = DifferenceOf(dataListPart1, x => x?.Cell14);

                // 平均值
                target.Cell15 = AverageOf(dataListPart1, x => x?.Cell15);

                // 差值
                target.Cell16 = DifferenceOf(dataListPart1, x => x?.Cell16);

                // 平均值
                target.Cell17 = AverageOf(dataListPart1, x => x?.Cell17);
                target.Cell18 = AverageOf(dataListPart1, x => x?.Cell18);
                target.Cell19 = AverageOf(dataListPart1, x => x?.Cell19);

                // 差值
                target.Cell20 = DifferenceOf(dataListPart1, x => x?.Cell20);

                // 平均值
                target.Cell21 = AverageOf(dataListPart1, x => x?.Cell21);
                target.Cell22 = AverageOf(dataListPart1, x => x?.Cell22);
                target.Cell23 = AverageOf(dataListPart1, x => x?.Cell23);

                // 最后一个值（保留原有逻辑，安全访问）
                target.Cell24 = LastValueOrNull(dataListPart1, x => x?.Cell24);

                // 平均值
                target.Cell25 = AverageOf(dataListPart1, x => x?.Cell25);
                target.Cell26 = AverageOf(dataListPart1, x => x?.Cell26);
                target.Cell27 = AverageOf(dataListPart1, x => x?.Cell27);
                target.Cell28 = AverageOf(dataListPart1, x => x?.Cell28);

                // 注释的人工填写字段（保留注释）
                //target.Cell29 = ...
                // ...

                // 平均值
                target.Cell36 = AverageOf(dataListPart1, x => x?.Cell36);

                // 差值
                target.Cell37 = DifferenceOf(dataListPart1, x => x?.Cell37);

                // 平均值
                target.Cell38 = AverageOf(dataListPart1, x => x?.Cell38);
                target.Cell39 = AverageOf(dataListPart1, x => x?.Cell39);
                target.Cell40 = AverageOf(dataListPart1, x => x?.Cell40);
                target.Cell41 = AverageOf(dataListPart1, x => x?.Cell41);

                // 差值
                target.Cell42 = DifferenceOf(dataListPart1, x => x?.Cell42);

                // 第二组：Cell51-Cell100
                target.Cell51 = AverageOf(dataListPart1, x => x?.Cell51);
                target.Cell52 = AverageOf(dataListPart1, x => x?.Cell52);
                target.Cell53 = AverageOf(dataListPart1, x => x?.Cell53);
                target.Cell54 = AverageOf(dataListPart1, x => x?.Cell54);

                // 差值
                target.Cell55 = DifferenceOf(dataListPart1, x => x?.Cell55);

                target.Cell61 = AverageOf(dataListPart1, x => x?.Cell61);
                target.Cell62 = AverageOf(dataListPart1, x => x?.Cell62);
                target.Cell63 = AverageOf(dataListPart1, x => x?.Cell63);
                target.Cell64 = AverageOf(dataListPart1, x => x?.Cell64);
                target.Cell65 = AverageOf(dataListPart1, x => x?.Cell65);
                target.Cell66 = AverageOf(dataListPart1, x => x?.Cell66);
                target.Cell67 = AverageOf(dataListPart1, x => x?.Cell67);
                target.Cell68 = AverageOf(dataListPart1, x => x?.Cell68);
                target.Cell69 = AverageOf(dataListPart1, x => x?.Cell69);
                target.Cell70 = AverageOf(dataListPart1, x => x?.Cell70);
                target.Cell71 = AverageOf(dataListPart1, x => x?.Cell71);
                target.Cell72 = AverageOf(dataListPart1, x => x?.Cell72);
                target.Cell73 = AverageOf(dataListPart1, x => x?.Cell73);
                target.Cell74 = AverageOf(dataListPart1, x => x?.Cell74);
                target.Cell75 = AverageOf(dataListPart1, x => x?.Cell75);
                target.Cell76 = AverageOf(dataListPart1, x => x?.Cell76);
                target.Cell77 = AverageOf(dataListPart1, x => x?.Cell77);
                target.Cell78 = AverageOf(dataListPart1, x => x?.Cell78);

                // 差值
                target.Cell79 = DifferenceOf(dataListPart1, x => x?.Cell79);

                target.Cell80 = AverageOf(dataListPart1, x => x?.Cell80);
                target.Cell81 = AverageOf(dataListPart1, x => x?.Cell81);

                target.Cell88 = AverageOf(dataListPart1, x => x?.Cell88);
                target.Cell89 = AverageOf(dataListPart1, x => x?.Cell89);
                target.Cell90 = AverageOf(dataListPart1, x => x?.Cell90);
                target.Cell91 = AverageOf(dataListPart1, x => x?.Cell91);
                target.Cell92 = AverageOf(dataListPart1, x => x?.Cell92);

                // 第三组：Cell101-Cell150
                target.Cell101 = AverageOf(dataListPart1, x => x?.Cell101);
                target.Cell102 = DifferenceOf(dataListPart1, x => x?.Cell102);
                target.Cell103 = AverageOf(dataListPart1, x => x?.Cell103);
                target.Cell104 = DifferenceOf(dataListPart1, x => x?.Cell104);

                target.Cell105 = AverageOf(dataListPart1, x => x?.Cell105);
                target.Cell106 = AverageOf(dataListPart1, x => x?.Cell106);
                target.Cell107 = AverageOf(dataListPart1, x => x?.Cell107);
                target.Cell108 = AverageOf(dataListPart1, x => x?.Cell108);
                target.Cell109 = AverageOf(dataListPart1, x => x?.Cell109);

                target.Cell110 = DifferenceOf(dataListPart1, x => x?.Cell110);

                target.Cell111 = AverageOf(dataListPart1, x => x?.Cell111);
                target.Cell112 = AverageOf(dataListPart1, x => x?.Cell112);
                target.Cell113 = AverageOf(dataListPart1, x => x?.Cell113);

                target.Cell114 = DifferenceOf(dataListPart1, x => x?.Cell114);

                target.Cell115 = AverageOf(dataListPart1, x => x?.Cell115);
                target.Cell116 = DifferenceOf(dataListPart1, x => x?.Cell116);
                target.Cell117 = AverageOf(dataListPart1, x => x?.Cell117);
                target.Cell118 = DifferenceOf(dataListPart1, x => x?.Cell118);

                target.Cell119 = AverageOf(dataListPart1, x => x?.Cell119);
                target.Cell120 = AverageOf(dataListPart1, x => x?.Cell120);
                target.Cell121 = AverageOf(dataListPart1, x => x?.Cell121);
                target.Cell122 = AverageOf(dataListPart1, x => x?.Cell122);
                target.Cell123 = AverageOf(dataListPart1, x => x?.Cell123);
                target.Cell124 = AverageOf(dataListPart1, x => x?.Cell124);
                target.Cell125 = AverageOf(dataListPart1, x => x?.Cell125);
                target.Cell126 = AverageOf(dataListPart1, x => x?.Cell126);
                target.Cell127 = AverageOf(dataListPart1, x => x?.Cell127);
                target.Cell128 = AverageOf(dataListPart1, x => x?.Cell128);

                target.Cell130 = AverageOf(dataListPart1, x => x?.Cell130);
                target.Cell131 = AverageOf(dataListPart1, x => x?.Cell131);

                target.Cell132 = DifferenceOf(dataListPart1, x => x?.Cell132);

                target.Cell133 = AverageOf(dataListPart1, x => x?.Cell133);
                target.Cell134 = AverageOf(dataListPart1, x => x?.Cell134);
            }
            List<OperatorInputData> rawDataPart2 = await _operatorInputData.GetByDateTimeRangeAsync(ReportInfo.TimeStart, ReportInfo.TimeEnd);
            List<OperatorInputData> dataListPart2 = SortDataByTime(rawDataPart2, ReportInfo.TimeStart);
            {
                // 使用相同的通用方法（OperatorInputData 类型相同字段名）
                target.Cell151 = AverageOf(dataListPart2, x => x?.Cell1);
                target.Cell152 = AverageOf(dataListPart2, x => x?.Cell2);
                target.Cell153 = AverageOf(dataListPart2, x => x?.Cell3);
                target.Cell154 = AverageOf(dataListPart2, x => x?.Cell4);
                target.Cell155 = AverageOf(dataListPart2, x => x?.Cell5);
                target.Cell156 = AverageOf(dataListPart2, x => x?.Cell6);
                target.Cell157 = AverageOf(dataListPart2, x => x?.Cell7);
                target.Cell158 = AverageOf(dataListPart2, x => x?.Cell8);
                target.Cell159 = AverageOf(dataListPart2, x => x?.Cell9);
                target.Cell160 = AverageOf(dataListPart2, x => x?.Cell10);
                target.Cell161 = AverageOf(dataListPart2, x => x?.Cell11);
                target.Cell162 = AverageOf(dataListPart2, x => x?.Cell12);
                target.Cell163 = AverageOf(dataListPart2, x => x?.Cell13);
                target.Cell164 = AverageOf(dataListPart2, x => x?.Cell14);
                target.Cell165 = AverageOf(dataListPart2, x => x?.Cell15);
                target.Cell166 = AverageOf(dataListPart2, x => x?.Cell16);
                target.Cell167 = AverageOf(dataListPart2, x => x?.Cell17);
                target.Cell168 = AverageOf(dataListPart2, x => x?.Cell18);
                target.Cell169 = AverageOf(dataListPart2, x => x?.Cell19);
                target.Cell170 = AverageOf(dataListPart2, x => x?.Cell20);
                target.Cell171 = AverageOf(dataListPart2, x => x?.Cell21);
                target.Cell172 = AverageOf(dataListPart2, x => x?.Cell22);
                target.Cell173 = AverageOf(dataListPart2, x => x?.Cell23);
                target.Cell174 = AverageOf(dataListPart2, x => x?.Cell24);
                target.Cell175 = AverageOf(dataListPart2, x => x?.Cell25);
                target.Cell176 = AverageOf(dataListPart2, x => x?.Cell26);
                target.Cell177 = AverageOf(dataListPart2, x => x?.Cell27);
                target.Cell178 = AverageOf(dataListPart2, x => x?.Cell28);
                target.Cell179 = AverageOf(dataListPart2, x => x?.Cell29);
                target.Cell80 = AverageOf(dataListPart2, x => x?.Cell30);
                target.Cell81 = AverageOf(dataListPart2, x => x?.Cell31);
                target.Cell82 = AverageOf(dataListPart2, x => x?.Cell32);
                target.Cell83 = AverageOf(dataListPart2, x => x?.Cell33);
                target.Cell84 = AverageOf(dataListPart2, x => x?.Cell34);
                target.Cell85 = AverageOf(dataListPart2, x => x?.Cell35);
                target.Cell86 = AverageOf(dataListPart2, x => x?.Cell36);
                target.Cell87 = AverageOf(dataListPart2, x => x?.Cell37);
                target.Cell88 = AverageOf(dataListPart2, x => x?.Cell38);
                target.Cell89 = AverageOf(dataListPart2, x => x?.Cell39);
                target.Cell90 = AverageOf(dataListPart2, x => x?.Cell40);
                target.Cell91 = AverageOf(dataListPart2, x => x?.Cell41);
                target.Cell92 = AverageOf(dataListPart2, x => x?.Cell42);
                target.Cell93 = AverageOf(dataListPart2, x => x?.Cell43);
                target.Cell94 = AverageOf(dataListPart2, x => x?.Cell44);
                target.Cell95 = AverageOf(dataListPart2, x => x?.Cell45);
                target.Cell96 = AverageOf(dataListPart2, x => x?.Cell46);
                target.Cell97 = AverageOf(dataListPart2, x => x?.Cell47);
                target.Cell98 = AverageOf(dataListPart2, x => x?.Cell48);
                target.Cell99 = AverageOf(dataListPart2, x => x?.Cell49);
                target.Cell200 = AverageOf(dataListPart2, x => x?.Cell50);
            }

            if (!existRecord)//无记录则插入
                await _calculatedDatas.AddAsync(target);
            await _reportUnitOfWork.SaveChangesAsync();
            return true;
        }

        private static async Task WeekDataCalculate(CalculatedData target, List<CalculatedData> dataListPart1)
        {
            target.PH = 80;//暂时没有特殊意义
            target.Cell1 = dataListPart1.Select(x => x.Cell1 ?? 0).Average();//平均值（保留原逻辑）
        }
        private static async Task MonthDataCalculate(CalculatedData target, List<CalculatedData> dataListPart1)
        {
            target.PH = 80;//暂时没有特殊意义
            target.Cell1 = dataListPart1.Select(x => x.Cell1 ?? 0).Average();//平均值（保留原逻辑）
        }
        private static async Task YearDataCalculate(CalculatedData target, List<CalculatedData> dataListPart1)
        {
            target.PH = 80;//暂时没有特殊意义
            target.Cell1 = dataListPart1.Select(x => x.Cell1 ?? 0).Average();//平均值（保留原逻辑）
        }

        //SourceData 按照时间顺序排序，确保每个小时的数据在正确的位置上  共25个小时的数据
        private static List<SourceData> SortDataByTime(List<SourceData> sourceData, DateTime baseDate)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<SourceData>(new SourceData[25]);
            for (int i = 0; i < 25; i++)
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

        //OperatorInputData 按照时间顺序排序，确保每个小时的数据在正确的位置上  共25个小时的数据
        private static List<OperatorInputData> SortDataByTime(List<OperatorInputData> sourceData, DateTime baseDate)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<OperatorInputData>(new OperatorInputData[25]);
            for (int i = 0; i < 25; i++)
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

        // ---------- 辅助方法 ----------
        // 通用平均值：排除 null 项；若无有效值返回 null（表示跳过）
        private static float? AverageOf<T>(IEnumerable<T> list, Func<T?, float?> selector)
            where T : class
        {
            if (list == null) return null;
            var vals = list
                .Where(x => x != null)
                .Select(x => selector(x!))
                .Where(v => v.HasValue)
                .Select(v => v!.Value);
            return vals.Any() ? (float?)vals.Average() : null;
        }

        // 通用差值（最后 - 第一个非空值），若无法计算则返回 null（表示跳过）
        private static float? DifferenceOf<T>(IEnumerable<T> list, Func<T?, float?> selector)
            where T : class
        {
            if (list == null) return null;
            var first = list.FirstOrDefault(x => x != null && selector(x!) != null);
            var last = list.LastOrDefault(x => x != null && selector(x!) != null);
            if (first != null && last != null)
            {
                var fv = selector(first);
                var lv = selector(last);
                if (fv.HasValue && lv.HasValue)
                    return lv.Value - fv.Value;
            }
            return null;
        }

        // 获取最后存在的某字段值，若不存在返回 null（用于保留“最后一个值”的原始逻辑但允许跳过）
        private static float? LastValueOrNull<T>(IEnumerable<T> list, Func<T?, float?> selector)
            where T : class
        {
            if (list == null) return null;
            var last = list.LastOrDefault(x => x != null && selector(x!) != null);
            return last != null ? selector(last) : null;
        }
    }
}
