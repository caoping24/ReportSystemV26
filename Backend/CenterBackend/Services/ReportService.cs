using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Reflection;

namespace CenterBackend.Services
{
    public class ReportService : IReportService
    {
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IReportRecordRepository<ReportRecord> _reportRecord;
        private readonly IReportRepository<CalculatedData> _calculatedDatas;
        private readonly IReportUnitOfWork _reportUnitOfWork;
        private readonly CenterReportDbContext _dbContext;
        // 构造函数注入：按顺序注入5个SourceData仓储 + 原有依赖，一一对应赋值
        public ReportService(IReportRepository<SourceData> SourceData,
                             IReportRecordRepository<ReportRecord> reportRecord,
                             IReportRepository<CalculatedData> CalculatedDatas,
                             IReportUnitOfWork reportUnitOfWork,
                             //IHttpContextAccessor httpContextAccessor,
                            CenterReportDbContext _dbContext)
        {
            this._sourceData = SourceData;
            this._reportRecord = reportRecord;
            this._calculatedDatas = CalculatedDatas;
            this._reportUnitOfWork = reportUnitOfWork;
            this._dbContext = _dbContext;
        }


        //public async Task<bool> DeleteReport(long id, DailyInsertDto _AddReportDailyDto)
        //{
        //    return true;
        //}

        //public async Task<bool> AddReport(DailyInsertDto _AddReportDailyDto)
        //{
        //    return true;
        //}

        /// <summary>
        /// 根据传入的Type类型，计算对应维度的统计数据并插入到CalculatedData表中 注意传入的时间
        /// </summary>
        /// <param name="_Dto"></param>
        /// <returns></returns>
        public async Task<bool> DataAnalyses(CalculateAndInsertDto _Dto)
        {
            DateTime StartTime;
            DateTime StopTime;

            switch (_Dto.Type)
            {
                case 1: // 昨天
                    StartTime = _Dto.Time.Date.AddDays(-1).AddHours(8); // 开始时间等于昨天的8点0分
                    StopTime = StartTime.AddHours(24).AddMinutes(59); // 结束时间等于今天的8点59分
                    break;
                case 2: // 上周
                    DateTime currentDayOfWeek = _Dto.Time.Date;// 计算上周的开始时间（星期一）
                    int daysToLastMonday = ((int)currentDayOfWeek.DayOfWeek + 6) % 7 + 7;
                    StartTime = currentDayOfWeek.AddDays(-daysToLastMonday);
                    StopTime = StartTime.AddDays(6).AddHours(23).AddMinutes(59);// 计算上周的结束时间（星期天）
                    break;
                case 3: // 上月
                    StartTime = new DateTime(_Dto.Time.Year, _Dto.Time.Month, 1).AddMonths(-1);// 计算上月的开始时间（1号）
                    StopTime = StartTime.AddMonths(1).AddDays(-1).AddHours(23).AddMinutes(59);// 计算上月的结束时间（最后一天）
                    break;
                case 4: // 去年   
                    StartTime = new DateTime(_Dto.Time.Year, 1, 1).AddYears(-1);// 计算去年的开始时间（1月1号）
                    StopTime = new DateTime(_Dto.Time.Year, 1, 1).AddDays(-1).AddHours(23).AddMinutes(59);// 计算去年的结束时间（12月31号）
                    break;
                default:
                    return false;
            }
            return await CalculatedDataAndInsert(StartTime, StopTime, _Dto.Type);
        }

        /// <summary>
        /// 根据Tpye类型，计算周/月/年统计数据
        /// </summary>
        private async Task<bool> CalculatedDataAndInsert(DateTime startTime, DateTime stopTime, int type)
        {
            var ReportedTime = startTime.Date;//记录是那一天的数据
            var target = _calculatedDatas.Db.FirstOrDefault(r => r.Type == type && r.ReportedTime == ReportedTime);
            bool isNewRecord = (target == null );
            if (isNewRecord)
            {
                target = new CalculatedData
                {
                    Type = type,
                    ReportedTime = ReportedTime,
                };
            }
            else
            {
                target.ReportedTime = DateTime.Now; // 更新时刷新创建时间（或改updateTime更合理）
            }

            bool isCalculatedSuccess = await CalculateDimensionDataAsync(target, startTime, stopTime, type);
            if (!isCalculatedSuccess)
            {
                return false;
            }

            if (isNewRecord)
            {
                await _calculatedDatas.AddAsync(target);
            }
            else
            {
                await _calculatedDatas.Update(target);

            }

            await _reportUnitOfWork.SaveChangesAsync();
            return true;
        }


        /// <summary>
        /// 提取：按维度计算数据（解耦维度计算逻辑，便于维护）
        /// </summary>
        /// <param name="target">要赋值的统计对象</param>
        /// <param name="startTime">开始时间</param>
        /// <param name="stopTime">结束时间</param>
        /// <param name="analysisType">统计维度</param>
        /// <returns>是否计算成功</returns>
        private async Task<bool> CalculateDimensionDataAsync(CalculatedData target, DateTime startTime, DateTime stopTime, int analysisType)
        {
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target), "统计目标对象不能为空");
            }

            switch (analysisType)
            {
                case 1:
                    var dailyDataList = await _sourceData.GetByDateTimeRangeAsync(startTime, stopTime);
                    if (dailyDataList.Count == 0) return false;
                    await DayDataCalculate(target, dailyDataList);
                    break;

                case 2:
                    var weeklyDataList = await _calculatedDatas.GetByDateTimeRangeAsync(startTime, stopTime, 1);
                    if (weeklyDataList.Count == 0) return false;
                    await WeekDataCalculate(target, weeklyDataList);
                    break;

                case 3:
                    var monthlyDataList = await _calculatedDatas.GetByDateTimeRangeAsync(startTime, stopTime, 1);
                    if (monthlyDataList.Count == 0) return false;
                    await MonthDataCalculate(target, monthlyDataList);
                    break;

                case 4:
                    var yearlyDataList = await _calculatedDatas.GetByDateTimeRangeAsync(startTime, stopTime, 3);
                    if (yearlyDataList.Count == 0) return false;
                    await YearDataCalculate(target, yearlyDataList);
                    break;

                default:
                    return false;
            }

            return true;
        }
        private static async Task DayDataCalculate(CalculatedData target, List<SourceData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义
            //(dataList.Last().Cell2 ?? 0) - (dataList.First().Cell2 ?? 0);//差值
            // dataList.Select(x => x.Cell3 ?? 0).Sum();//总和
            target.Cell1 = dataList.Select(x => x.Cell1 ?? 0).Average();//平均值
            target.Cell2 = dataList.Select(x => x.Cell2 ?? 0).Average();
            target.Cell3 = dataList.Select(x => x.Cell3 ?? 0).Average();
            target.Cell4 = dataList.Last().Cell4 - dataList.First().Cell4;//差值
            target.Cell5 = dataList.Last().Cell5 - dataList.First().Cell5;//差值
            target.Cell6 = dataList.Select(x => x.Cell6 ?? 0).Average();
            target.Cell7 = dataList.Select(x => x.Cell7 ?? 0).Average();
            target.Cell8 = dataList.Last().Cell8 - dataList.First().Cell8;//差值
            target.Cell9 = dataList.Select(x => x.Cell9 ?? 0).Average();
            target.Cell10 = dataList.Select(x => x.Cell10 ?? 0).Average();
            target.Cell11 = dataList.Select(x => x.Cell11 ?? 0).Average();
            target.Cell12 = dataList.Select(x => x.Cell12 ?? 0).Average();
            target.Cell13 = dataList.Select(x => x.Cell13 ?? 0).Average();
            target.Cell14 = dataList.Last().Cell14 - dataList.First().Cell14;//差值
            target.Cell15 = dataList.Select(x => x.Cell15 ?? 0).Average();
            target.Cell16 = dataList.Last().Cell16 - dataList.First().Cell16;//差值
            target.Cell17 = dataList.Select(x => x.Cell17 ?? 0).Average();
            target.Cell18 = dataList.Select(x => x.Cell18 ?? 0).Average();
            target.Cell19 = dataList.Select(x => x.Cell19 ?? 0).Average();
            target.Cell20 = dataList.Last().Cell20 - dataList.First().Cell20;//差值
            target.Cell21 = dataList.Select(x => x.Cell21 ?? 0).Average();
            target.Cell22 = dataList.Select(x => x.Cell22 ?? 0).Average();
            target.Cell23 = dataList.Select(x => x.Cell23 ?? 0).Average();
            target.Cell24 = dataList.Last().Cell24;//最后一个值
            target.Cell25 = dataList.Select(x => x.Cell25 ?? 0).Average();
            target.Cell26 = dataList.Select(x => x.Cell26 ?? 0).Average();
            target.Cell27 = dataList.Select(x => x.Cell27 ?? 0).Average();
            target.Cell28 = dataList.Select(x => x.Cell28 ?? 0).Average();
            //target.Cell29 = dataList.Select(x => x.Cell29 ?? 0).Average();//人工填写
            //target.Cell30 = dataList.Select(x => x.Cell30 ?? 0).Average();//人工填写
            //target.Cell31 = dataList.Select(x => x.Cell31 ?? 0).Average();//人工填写
            //target.Cell32 = dataList.Select(x => x.Cell32 ?? 0).Average();//人工填写
            //target.Cell33 = dataList.Select(x => x.Cell33 ?? 0).Average();//人工填写
            //target.Cell34 = dataList.Select(x => x.Cell34 ?? 0).Average();//人工填写
            //target.Cell35 = dataList.Select(x => x.Cell35 ?? 0).Average();//人工填写
            target.Cell36 = dataList.Select(x => x.Cell36 ?? 0).Average();
            target.Cell37 = dataList.Last().Cell37 - dataList.First().Cell37;//差值
            target.Cell38 = dataList.Select(x => x.Cell38 ?? 0).Average();
            target.Cell39 = dataList.Select(x => x.Cell39 ?? 0).Average();
            target.Cell40 = dataList.Select(x => x.Cell40 ?? 0).Average();
            target.Cell41 = dataList.Select(x => x.Cell41 ?? 0).Average();
            target.Cell42 = dataList.Last().Cell42 - dataList.First().Cell42;//差值
            //target.Cell43 = dataList.Select(x => x.Cell43 ?? 0).Average();
            //target.Cell44 = dataList.Select(x => x.Cell44 ?? 0).Average();
            //target.Cell45 = dataList.Select(x => x.Cell45 ?? 0).Average();
            //target.Cell46 = dataList.Select(x => x.Cell46 ?? 0).Average();
            //target.Cell47 = dataList.Select(x => x.Cell47 ?? 0).Average();
            //target.Cell48 = dataList.Select(x => x.Cell48 ?? 0).Average();
            //target.Cell49 = dataList.Select(x => x.Cell49 ?? 0).Average();
            //target.Cell50 = dataList.Select(x => x.Cell50 ?? 0).Average();
            // 第二组：Cell51-Cell100
            target.Cell51 = dataList.Select(x => x.Cell51 ?? 0).Average();
            target.Cell52 = dataList.Select(x => x.Cell52 ?? 0).Average();
            target.Cell53 = dataList.Select(x => x.Cell53 ?? 0).Average();
            target.Cell54 = dataList.Select(x => x.Cell54 ?? 0).Average();
            target.Cell55 = dataList.Last().Cell55 - dataList.First().Cell55;//差值
            //target.Cell56 = dataList.Select(x => x.Cell56 ?? 0).Average();
            //target.Cell57 = dataList.Select(x => x.Cell57 ?? 0).Average();
            //target.Cell58 = dataList.Select(x => x.Cell58 ?? 0).Average();
            //target.Cell59 = dataList.Select(x => x.Cell59 ?? 0).Average();
            //target.Cell60 = dataList.Select(x => x.Cell60 ?? 0).Average();
            target.Cell61 = dataList.Select(x => x.Cell61 ?? 0).Average();
            target.Cell62 = dataList.Select(x => x.Cell62 ?? 0).Average();
            target.Cell63 = dataList.Select(x => x.Cell63 ?? 0).Average();
            target.Cell64 = dataList.Select(x => x.Cell64 ?? 0).Average();
            target.Cell65 = dataList.Select(x => x.Cell65 ?? 0).Average();
            target.Cell66 = dataList.Select(x => x.Cell66 ?? 0).Average();
            target.Cell67 = dataList.Select(x => x.Cell67 ?? 0).Average();
            target.Cell68 = dataList.Select(x => x.Cell68 ?? 0).Average();
            target.Cell69 = dataList.Select(x => x.Cell69 ?? 0).Average();
            target.Cell70 = dataList.Select(x => x.Cell70 ?? 0).Average();
            target.Cell71 = dataList.Select(x => x.Cell71 ?? 0).Average();
            target.Cell72 = dataList.Select(x => x.Cell72 ?? 0).Average();
            target.Cell73 = dataList.Select(x => x.Cell73 ?? 0).Average();
            target.Cell74 = dataList.Select(x => x.Cell74 ?? 0).Average();
            target.Cell75 = dataList.Select(x => x.Cell75 ?? 0).Average();
            target.Cell76 = dataList.Select(x => x.Cell76 ?? 0).Average();
            target.Cell77 = dataList.Select(x => x.Cell77 ?? 0).Average();
            target.Cell78 = dataList.Select(x => x.Cell78 ?? 0).Average();
            target.Cell79 = dataList.Last().Cell79 - dataList.First().Cell79;//差值
            target.Cell80 = dataList.Select(x => x.Cell80 ?? 0).Average();
            target.Cell81 = dataList.Select(x => x.Cell81 ?? 0).Average();
            //target.Cell82 = dataList.Select(x => x.Cell82 ?? 0).Average();
            //target.Cell83 = dataList.Select(x => x.Cell83 ?? 0).Average();
            //target.Cell84 = dataList.Select(x => x.Cell84 ?? 0).Average();
            //target.Cell85 = dataList.Select(x => x.Cell85 ?? 0).Average();
            //target.Cell86 = dataList.Select(x => x.Cell86 ?? 0).Average();
            //target.Cell87 = dataList.Select(x => x.Cell87 ?? 0).Average();
            target.Cell88 = dataList.Select(x => x.Cell88 ?? 0).Average();
            target.Cell89 = dataList.Select(x => x.Cell89 ?? 0).Average();
            target.Cell90 = dataList.Select(x => x.Cell90 ?? 0).Average();
            target.Cell91 = dataList.Select(x => x.Cell91 ?? 0).Average();
            target.Cell92 = dataList.Select(x => x.Cell92 ?? 0).Average();
            //target.Cell93 = dataList.Select(x => x.Cell93 ?? 0).Average();
            //target.Cell94 = dataList.Select(x => x.Cell94 ?? 0).Average();
            //target.Cell95 = dataList.Select(x => x.Cell95 ?? 0).Average();
            //target.Cell96 = dataList.Select(x => x.Cell96 ?? 0).Average();
            //target.Cell97 = dataList.Select(x => x.Cell97 ?? 0).Average();
            //target.Cell98 = dataList.Select(x => x.Cell98 ?? 0).Average();
            //target.Cell99 = dataList.Select(x => x.Cell99 ?? 0).Average();
            //target.Cell100 = dataList.Select(x => x.Cell100 ?? 0).Average();
            // 第三组：Cell101-Cell150
            target.Cell101 = dataList.Select(x => x.Cell101 ?? 0).Average();
            target.Cell102 = dataList.Last().Cell102 - dataList.First().Cell102;//差值
            target.Cell103 = dataList.Select(x => x.Cell103 ?? 0).Average();
            target.Cell104 = dataList.Last().Cell104 - dataList.First().Cell104;//差值
            target.Cell105 = dataList.Select(x => x.Cell105 ?? 0).Average();
            target.Cell106 = dataList.Select(x => x.Cell106 ?? 0).Average();
            target.Cell107 = dataList.Select(x => x.Cell107 ?? 0).Average();
            target.Cell108 = dataList.Select(x => x.Cell108 ?? 0).Average();
            target.Cell109 = dataList.Select(x => x.Cell109 ?? 0).Average();
            target.Cell110 = dataList.Last().Cell110 - dataList.First().Cell110;//差值
            target.Cell111 = dataList.Select(x => x.Cell111 ?? 0).Average();
            target.Cell112 = dataList.Select(x => x.Cell112 ?? 0).Average();
            target.Cell113 = dataList.Select(x => x.Cell113 ?? 0).Average();
            target.Cell114 = dataList.Last().Cell114 - dataList.First().Cell114;//差值
            target.Cell115 = dataList.Select(x => x.Cell115 ?? 0).Average();
            target.Cell116 = dataList.Last().Cell116 - dataList.First().Cell116;//差值
            target.Cell117 = dataList.Select(x => x.Cell117 ?? 0).Average();
            target.Cell118 = dataList.Last().Cell118 - dataList.First().Cell118;//差值
            target.Cell119 = dataList.Select(x => x.Cell119 ?? 0).Average();
            target.Cell120 = dataList.Select(x => x.Cell120 ?? 0).Average();
            target.Cell121 = dataList.Select(x => x.Cell121 ?? 0).Average();
            target.Cell122 = dataList.Select(x => x.Cell122 ?? 0).Average();
            target.Cell123 = dataList.Select(x => x.Cell123 ?? 0).Average();
            target.Cell124 = dataList.Select(x => x.Cell124 ?? 0).Average();
            target.Cell125 = dataList.Select(x => x.Cell125 ?? 0).Average();
            target.Cell126 = dataList.Select(x => x.Cell126 ?? 0).Average();
            target.Cell127 = dataList.Select(x => x.Cell127 ?? 0).Average();
            target.Cell128 = dataList.Select(x => x.Cell128 ?? 0).Average();
            //target.Cell129 = dataList.Last().Cell129 - dataList.First().Cell129;//差值
            target.Cell130 = dataList.Select(x => x.Cell130 ?? 0).Average();
            target.Cell131 = dataList.Select(x => x.Cell131 ?? 0).Average();
            target.Cell132 = dataList.Last().Cell132 - dataList.First().Cell132;//差值
            target.Cell133 = dataList.Select(x => x.Cell133 ?? 0).Average();
            target.Cell134 = dataList.Select(x => x.Cell134 ?? 0).Average();
            //target.Cell135 = dataList.Select(x => x.Cell135 ?? 0).Average();
            //target.Cell136 = dataList.Select(x => x.Cell136 ?? 0).Average();
            //target.Cell137 = dataList.Select(x => x.Cell137 ?? 0).Average();
            //target.Cell138 = dataList.Select(x => x.Cell138 ?? 0).Average();
            //target.Cell139 = dataList.Select(x => x.Cell139 ?? 0).Average();
            //target.Cell140 = dataList.Select(x => x.Cell140 ?? 0).Average();
            //target.Cell141 = dataList.Select(x => x.Cell141 ?? 0).Average();
            //target.Cell142 = dataList.Select(x => x.Cell142 ?? 0).Average();
            //target.Cell143 = dataList.Select(x => x.Cell143 ?? 0).Average();
            //target.Cell144 = dataList.Select(x => x.Cell144 ?? 0).Average();
            //target.Cell145 = dataList.Select(x => x.Cell145 ?? 0).Average();
            //target.Cell146 = dataList.Select(x => x.Cell146 ?? 0).Average();
            //target.Cell147 = dataList.Select(x => x.Cell147 ?? 0).Average();
            //target.Cell148 = dataList.Select(x => x.Cell148 ?? 0).Average();
            //target.Cell149 = dataList.Select(x => x.Cell149 ?? 0).Average();
            //target.Cell150 = dataList.Select(x => x.Cell150 ?? 0).Average();

        }
        private static async Task WeekDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.Cell1 = dataList.Select(x => x.Cell1 ?? 0).Average();//平均值
        }
        private static async Task MonthDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.Cell1 = dataList.Select(x => x.Cell1 ?? 0).Average();//平均值
        }
        private static async Task YearDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.Cell1 = dataList.Select(x => x.Cell1 ?? 0).Average();//平均值
        }

        /// <summary>
        /// 从模板文件创建文件流，然后按区域写数据并且保存到本地文件 Daily
        /// </summary>
        /// <param name="ModelFullPath">模板完整路径</param>
        /// <param name="TargetPullPath">生成文件的保存路径</param>
        /// <param name="ReportTime">报表日期</param>
        /// <returns></returns>
        public async Task<IActionResult> WriteXlsxAndSave(string ModelFullPath, string TargetPullPath, DateTime ReportTime, int Type)
        {
            DateTime StartTime;
            DateTime StopTime;
            List<SourceData>? dataList;
            List<CalculatedData>? dataList2;
            DateTime TempRepoetName = DateTime.Parse("2026-01-01");
            try
            {
                using var templateStream = new FileStream(ModelFullPath, FileMode.Open, FileAccess.Read);
                using var workbook = new XSSFWorkbook(templateStream);

                switch (Type)
                {
                    case 1: //本日

                        //if (ReportTime.AddMinutes(-1).Hour < 8)
                        //{
                        //    return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 应大于8:00" });
                        //}
                        dataList = await _sourceData.GetByDateTimeRangeAsync(ReportTime, ReportTime);
                        if (dataList == null || dataList.Count == 0)
                        {
                            return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 无数据" });
                        }
                        SourceData?[] targetArray = MatchSourceDataDay(dataList, ReportTime);

                        WriteXlsxDaily1(workbook, targetArray, ReportTime);
                        WriteXlsxDaily2(workbook, targetArray, ReportTime);
                        TempRepoetName = ReportTime.Date;
                        break;
                    case 2: // 本周
                        DateTime currentDayOfWeek = ReportTime.Date;// 计算本周一
                        int daysToThisMonday = ((int)currentDayOfWeek.DayOfWeek + 6) % 7;
                        StartTime = currentDayOfWeek.AddDays(-daysToThisMonday);
                        StopTime = StartTime.AddDays(6).AddHours(23).AddMinutes(59);

                        dataList2 = await _calculatedDatas.GetByDateTimeRangeAsync(StartTime, StopTime, 2);
                        if (dataList2 == null || dataList2.Count == 0)
                        {
                            return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 无数据" });
                        }
                        CalculatedData?[] targetArray2 = MatchSourceDataWeek(dataList2, StartTime);
                        WriteXlsxWeekly1(workbook, targetArray2, StartTime);
                        TempRepoetName = StartTime.Date;
                        break;
                    case 3: // 本月
                        StartTime = ReportTime.Date;// 计算上月的开始时间（1号）
                        StopTime = StartTime.AddMonths(1).AddDays(-1);
                        dataList2 = await _calculatedDatas.GetByDateTimeRangeAsync(StartTime, StopTime, 3);
                        if (dataList2 == null || dataList2.Count == 0)
                        {
                            return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 无数据" });
                        }
                        targetArray2 = MatchSourceDataMonth(dataList2, StartTime);
                        WriteXlsxMonthly(workbook, targetArray2, ReportTime);
                        TempRepoetName = StartTime.Date;
                        break;
                    case 4: // 今年   
                        StartTime = ReportTime.Date;
                        StopTime = StartTime.AddYears(1).AddMinutes(-1);
                        dataList2 = await _calculatedDatas.GetByDateTimeRangeAsync(StartTime, StopTime, 4);
                        if (dataList2 == null || dataList2.Count == 0)
                        {
                            return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 无数据" });
                        }
                        targetArray2 = MatchSourceDataYear(dataList2, StartTime);
                        WriteXlsxYearly(workbook, targetArray2, ReportTime);
                        TempRepoetName = StartTime.Date;
                        break;
                    default:
                        return new OkObjectResult(new { success = false, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} 类型无效" });
                }

                using var outputStream = new FileStream(TargetPullPath, FileMode.Create, FileAccess.Write);// 保存文件到指定路径
                workbook.Write(outputStream);

                //插入build记录(只更新时间)
                var existingRecord = _reportRecord.db.FirstOrDefault(r => r.Type == Type && r.ReportedTime == TempRepoetName);
                if (existingRecord == null)//无记录则新增，有记录则更新
                {
                    var temp = new ReportRecord();
                    temp.Type = Type;
                    temp.ReportedTime = TempRepoetName;
                    temp.ReportedTime = DateTime.Now;
                    await _reportRecord.AddAsync(temp);
                }
                else
                {
                    existingRecord.ReportedTime = DateTime.Now;
                }
                await _reportUnitOfWork.SaveChangesAsync();
                return new OkObjectResult(new { success = true, msg = $"类型:{Type} 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} Excel生成成功" });
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"生成Excel异常:类型:{Type}, 时间:{ReportTime:yyyy-MM-dd hh:mm:ss} ，异常信息：{ex}" });
            }
        }

        /// <summary>
        /// 适配SourceData实体：按小时差匹配到指定长度数组 存档日报
        /// </summary>
        /// <param name="sourceDataList">SourceData原始数据</param>
        /// <param name="startTime">时间起始点</param>
        /// <returns>指定长度的SourceData?[]</returns>
        public SourceData?[] MatchSourceDataDay(List<SourceData> sourceDataList, DateTime startTime)
        {
            SourceData?[] targetArray = new SourceData?[25];
            if (sourceDataList == null || sourceDataList.Count == 0) return targetArray;
            startTime = startTime.Date.AddHours(8);
            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.ReportedTime - startTime).TotalHours;
                int matchIndex = (int)Math.Round(hourDiff);
                if (matchIndex >= 0 && matchIndex < 25)
                {
                    targetArray[matchIndex] = data;
                }
            }
            return targetArray;
        }

        /// <summary>
        /// 适配SourceData实体：按小时差匹配到指定长度数组 存档Week
        /// </summary>
        /// <param name="sourceDataList">SourceData原始数据</param>
        /// <param name="startTime">时间起始点</param>
        /// <returns>指定长度的SourceData?[]</returns>
        public CalculatedData?[] MatchSourceDataWeek(List<CalculatedData> sourceDataList, DateTime startTime)
        {
            CalculatedData?[] targetArray = new CalculatedData?[8];
            if (sourceDataList == null || sourceDataList.Count == 0) return targetArray;

            startTime = startTime.AddHours(8); //上周一8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.ReportedTime - startTime).TotalDays;
                int matchIndex = (int)Math.Round(hourDiff);
                if (matchIndex >= 0 && matchIndex < 8)
                {
                    targetArray[matchIndex] = data;
                }
            }
            return targetArray;
        }
        /// <summary>
        /// 适配SourceData实体：按小时差匹配到指定长度数组 存档month
        /// </summary>
        /// <param name="sourceDataList">SourceData原始数据</param>
        /// <param name="startTime">时间起始点</param>
        /// <returns>指定长度的SourceData?[]</returns>
        public CalculatedData?[] MatchSourceDataMonth(List<CalculatedData> sourceDataList, DateTime startTime)
        {
            CalculatedData?[] targetArray = new CalculatedData?[32];
            if (sourceDataList == null || sourceDataList.Count == 0) return targetArray;

            startTime = startTime.AddHours(8); //上月1号8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.ReportedTime - startTime).TotalDays;
                int matchIndex = (int)Math.Round(hourDiff);
                if (matchIndex >= 0 && matchIndex < 32)
                {
                    targetArray[matchIndex] = data;
                }
            }
            return targetArray;
        }
        /// <summary>
        /// 适配SourceData实体：按小时差匹配到指定长度数组 存档year
        /// </summary>
        /// <param name="sourceDataList">SourceData原始数据</param>
        /// <param name="startTime">时间起始点</param>
        /// <returns>指定长度的SourceData?[]</returns>
        public CalculatedData?[] MatchSourceDataYear(List<CalculatedData> sourceDataList, DateTime startTime)
        {
            CalculatedData?[] targetArray = new CalculatedData?[13];
            if (sourceDataList == null || sourceDataList.Count == 0) return targetArray;

            startTime = startTime.AddHours(8); //去年1月1号8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;

                int matchIndex = (data.ReportedTime.Year - startTime.Year) * 12 + (data.ReportedTime.Month - startTime.Month);
                if (matchIndex >= 0 && matchIndex < 13)
                {
                    targetArray[matchIndex] = data;
                }
            }
            return targetArray;
        }
        /// <summary>
        /// 写Xlsx数据  白班
        /// </summary>
        private static bool WriteXlsxDaily1(XSSFWorkbook srcWorkbook, SourceData?[] dataList, DateTime ReportDataTime)
        {
            ISheet srcSheet = srcWorkbook.GetSheetAt(0); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy-MM-dd");
            SetXlsxCellString(srcSheet, 51, 1, Temp);//记录日期,昨日
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 0; i < 13; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过

                int Range1 = 5 + i;
                int Range2 = 21 + i;
                int Range3 = 38 + i;

                // 从Excel第2列开始写入
                //Rang1 
                if (data.Cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.Cell1.Value); }
                if (data.Cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.Cell2.Value); }
                if (data.Cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.Cell3.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell4 != null && prevData != null && prevData != null && prevData.Cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell4);
                        float prevVal = Convert.ToSingle(prevData.Cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell5 != null && prevData != null && prevData.Cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell5);
                        float prevVal = Convert.ToSingle(prevData.Cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.Cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.Cell6.Value); }
                if (data.Cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.Cell7.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell8 != null && prevData != null && prevData.Cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell8);
                        float prevVal = Convert.ToSingle(prevData.Cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.Cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.Cell9.Value); }
                if (data.Cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.Cell10.Value); }
                if (data.Cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.Cell11.Value); }
                if (data.Cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.Cell12.Value); }
                if (data.Cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.Cell13.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell14 != null && prevData != null && prevData.Cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell14);
                        float prevVal = Convert.ToSingle(prevData.Cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.Cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.Cell15.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell16 != null && prevData != null && prevData.Cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell16);
                        float prevVal = Convert.ToSingle(prevData.Cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.Cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.Cell17.Value); }
                if (data.Cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.Cell18.Value); }
                if (data.Cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.Cell19.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell20 != null && prevData != null && prevData.Cell20 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell20);
                        float prevVal = Convert.ToSingle(prevData.Cell20);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 21, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.Cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.Cell21.Value); }
                if (data.Cell22 != null) //摩尔比 大于0 小于2 
                {
                    if (data.Cell22.Value < 2)
                        SetXlsxCellValue(srcSheet, Range1, 23, data.Cell22.Value);
                }
                if (i == 12)
                {
                    if (data.Cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.Cell23.Value); }//只记录最后一个值
                }
                if (data.Cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.Cell24.Value); }

                if (data.Cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.Cell25.Value); }
                if (data.Cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.Cell26.Value); }
                if (data.Cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.Cell27.Value); }
                if (data.Cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.Cell28.Value); }
                //人工检测数据
                if (data.Cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.Cell29.Value); }
                if (data.Cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.Cell30.Value); }
                if (data.Cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.Cell31.Value); }
                if (data.Cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.Cell32.Value); }
                if (data.Cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.Cell33.Value); }
                if (data.Cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.Cell34.Value); }
                if (data.Cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.Cell35.Value); }

                if (data.Cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.Cell36.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell37 != null && prevData != null && prevData.Cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell37);
                        float prevVal = Convert.ToSingle(prevData.Cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.Cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.Cell38.Value); }
                if (data.Cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.Cell39.Value); }
                if (data.Cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.Cell40.Value); }
                if (data.Cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.Cell41.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell42 != null && prevData != null && prevData.Cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell42);
                        float prevVal = Convert.ToSingle(prevData.Cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.Cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.Cell43.Value); }
                //if (data.Cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.Cell44.Value); }
                //if (data.Cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.Cell45.Value); }
                //if (data.Cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.Cell46.Value); }
                //if (data.Cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.Cell47.Value); }
                //if (data.Cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.Cell48.Value); }
                //if (data.Cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.Cell49.Value); }
                //if (data.Cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.Cell50.Value); }

                //Rang2
                if (data.Cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.Cell51.Value); }
                if (data.Cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.Cell52.Value); }
                if (data.Cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.Cell53.Value); }
                if (data.Cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.Cell54.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell55 != null && prevData != null && prevData.Cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell55);
                        float prevVal = Convert.ToSingle(prevData.Cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                //人工检测数据
                if (data.Cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.Cell56.Value); }
                if (data.Cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.Cell57.Value); }
                if (data.Cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.Cell58.Value); }
                if (data.Cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.Cell59.Value); }
                if (data.Cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.Cell60.Value); }

                if (data.Cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.Cell61.Value); }
                if (data.Cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.Cell62.Value); }
                if (data.Cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.Cell63.Value); }
                if (data.Cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.Cell64.Value); }
                if (data.Cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.Cell65.Value); }
                if (data.Cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.Cell66.Value); }
                if (data.Cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.Cell67.Value); }
                if (data.Cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.Cell68.Value); }
                if (data.Cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.Cell69.Value); }
                if (data.Cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.Cell70.Value); }
                if (data.Cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.Cell71.Value); }
                if (data.Cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.Cell72.Value); }
                if (data.Cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.Cell73.Value); }
                if (data.Cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.Cell74.Value); }
                if (data.Cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.Cell75.Value); }
                if (data.Cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.Cell76.Value); }
                if (data.Cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.Cell77.Value); }
                if (data.Cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.Cell78.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell79 != null && prevData != null && prevData.Cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell79);
                        float prevVal = Convert.ToSingle(prevData.Cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.Cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.Cell80.Value); }
                if (data.Cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.Cell81.Value); }
                //人工检测数据
                if (data.Cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.Cell82.Value); }
                if (data.Cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.Cell83.Value); }
                if (data.Cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.Cell84.Value); }
                if (data.Cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.Cell85.Value); }
                if (data.Cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.Cell86.Value); }
                if (data.Cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.Cell87.Value); }

                if (data.Cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.Cell88.Value); }
                if (data.Cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.Cell89.Value); }
                if (data.Cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.Cell90.Value); }
                if (data.Cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.Cell91.Value); }
                if (data.Cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.Cell92.Value); }
                //if (data.Cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.Cell93.Value); }
                //if (data.Cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.Cell94.Value); }
                //if (data.Cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.Cell95.Value); }
                //if (data.Cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.Cell96.Value); }
                //if (data.Cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.Cell97.Value); }
                //if (data.Cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.Cell98.Value); }
                //if (data.Cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.Cell99.Value); }
                //if (data.Cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.Cell100.Value); }

                //Rang3
                if (data.Cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.Cell101.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell102 != null && prevData != null && prevData.Cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell102);
                        float prevVal = Convert.ToSingle(prevData.Cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.Cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.Cell103.Value); }
                if (data.Cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.Cell104.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell105 != null && prevData != null && prevData.Cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell105);
                        float prevVal = Convert.ToSingle(prevData.Cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.Cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.Cell106.Value); }
                if (data.Cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.Cell107.Value); }
                if (data.Cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.Cell108.Value); }
                if (data.Cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.Cell109.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell110 != null && prevData != null && prevData.Cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell110);
                        float prevVal = Convert.ToSingle(prevData.Cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.Cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.Cell111.Value); }
                if (data.Cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.Cell112.Value); }
                if (data.Cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.Cell113.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell114 != null && prevData != null && prevData.Cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell114);
                        float prevVal = Convert.ToSingle(prevData.Cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.Cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.Cell115.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell116 != null && prevData != null && prevData.Cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell116);
                        float prevVal = Convert.ToSingle(prevData.Cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.Cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.Cell117.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell118 != null && prevData != null && prevData.Cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell118);
                        float prevVal = Convert.ToSingle(prevData.Cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.Cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.Cell119.Value); }
                if (data.Cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.Cell120.Value); }
                if (data.Cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.Cell121.Value); }
                if (data.Cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.Cell122.Value); }
                if (data.Cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.Cell123.Value); }
                if (data.Cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.Cell124.Value); }
                if (data.Cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.Cell125.Value); }
                if (data.Cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.Cell126.Value); }
                if (data.Cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.Cell127.Value); }
                if (data.Cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.Cell128.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell129 != null && prevData != null && prevData.Cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell129);
                        float prevVal = Convert.ToSingle(prevData.Cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.Cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.Cell130.Value); }
                if (data.Cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.Cell131.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell132 != null && prevData != null && prevData.Cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell132);
                        float prevVal = Convert.ToSingle(prevData.Cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.Cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.Cell133.Value); }
                if (data.Cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.Cell134.Value); }
                //人工检测数据
                if (data.Cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.Cell135.Value); }
                if (data.Cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.Cell136.Value); }
                if (data.Cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.Cell137.Value); }
                if (data.Cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.Cell138.Value); }
                if (data.Cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.Cell139.Value); }
                if (data.Cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.Cell140.Value); }
                if (data.Cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.Cell141.Value); }

                //if (data.Cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.Cell142.Value); }
                //if (data.Cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.Cell143.Value); }
                //if (data.Cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.Cell144.Value); }
                //if (data.Cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.Cell145.Value); }
                //if (data.Cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.Cell146.Value); }
                //if (data.Cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.Cell147.Value); }
                //if (data.Cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.Cell148.Value); }
                //if (data.Cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.Cell149.Value); }
                //if (data.Cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.Cell150.Value); }

            }
            return true;
        }

        /// <summary>
        /// 写Xlsx数据  夜班
        /// </summary>
        private static bool WriteXlsxDaily2(XSSFWorkbook srcWorkbook, SourceData?[] dataList, DateTime ReportDataTime)
        {

            ISheet srcSheet = srcWorkbook.GetSheetAt(2); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy-MM-dd");
            SetXlsxCellString(srcSheet, 51, 1, Temp);//记录日期
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 12; i < 25; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过

                int Range1 = 5 + i - 12;
                int Range2 = 21 + i - 12;
                int Range3 = 38 + i - 12;

                // 从Excel第2列开始写入
                //Rang1 
                if (data.Cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.Cell1.Value); }
                if (data.Cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.Cell2.Value); }
                if (data.Cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.Cell3.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell4 != null && prevData != null && prevData.Cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell4);
                        float prevVal = Convert.ToSingle(prevData.Cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell5 != null && prevData != null && prevData.Cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell5);
                        float prevVal = Convert.ToSingle(prevData.Cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.Cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.Cell6.Value); }
                if (data.Cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.Cell7.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell8 != null && prevData != null && prevData.Cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell8);
                        float prevVal = Convert.ToSingle(prevData.Cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.Cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.Cell9.Value); }
                if (data.Cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.Cell10.Value); }
                if (data.Cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.Cell11.Value); }
                if (data.Cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.Cell12.Value); }
                if (data.Cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.Cell13.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell14 != null && prevData != null && prevData.Cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell14);
                        float prevVal = Convert.ToSingle(prevData.Cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.Cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.Cell15.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell16 != null && prevData != null && prevData.Cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell16);
                        float prevVal = Convert.ToSingle(prevData.Cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.Cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.Cell17.Value); }
                if (data.Cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.Cell18.Value); }
                if (data.Cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.Cell19.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell20 != null && prevData != null && prevData.Cell20 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell20);
                        float prevVal = Convert.ToSingle(prevData.Cell20);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 21, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.Cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.Cell21.Value); }
                if (data.Cell22 != null) //摩尔比 大于0 小于2 
                {
                    if (data.Cell22.Value < 2)
                        SetXlsxCellValue(srcSheet, Range1, 23, data.Cell22.Value);
                }
                if (i == 24)
                {
                    if (data.Cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.Cell23.Value); }//只记录最后一个值
                }
                if (data.Cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.Cell24.Value); }

                if (data.Cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.Cell25.Value); }
                if (data.Cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.Cell26.Value); }
                if (data.Cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.Cell27.Value); }
                if (data.Cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.Cell28.Value); }
                //人工检测数据
                if (data.Cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.Cell29.Value); }
                if (data.Cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.Cell30.Value); }
                if (data.Cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.Cell31.Value); }
                if (data.Cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.Cell32.Value); }
                if (data.Cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.Cell33.Value); }
                if (data.Cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.Cell34.Value); }
                if (data.Cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.Cell35.Value); }

                if (data.Cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.Cell36.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell37 != null && prevData != null && prevData.Cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell37);
                        float prevVal = Convert.ToSingle(prevData.Cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.Cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.Cell38.Value); }
                if (data.Cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.Cell39.Value); }
                if (data.Cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.Cell40.Value); }
                if (data.Cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.Cell41.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell42 != null && prevData != null && prevData.Cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell42);
                        float prevVal = Convert.ToSingle(prevData.Cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.Cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.Cell43.Value); }
                //if (data.Cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.Cell44.Value); }
                //if (data.Cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.Cell45.Value); }
                //if (data.Cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.Cell46.Value); }
                //if (data.Cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.Cell47.Value); }
                //if (data.Cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.Cell48.Value); }
                //if (data.Cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.Cell49.Value); }
                //if (data.Cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.Cell50.Value); }

                //Rang2
                if (data.Cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.Cell51.Value); }
                if (data.Cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.Cell52.Value); }
                if (data.Cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.Cell53.Value); }
                if (data.Cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.Cell54.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell55 != null && prevData != null && prevData.Cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell55);
                        float prevVal = Convert.ToSingle(prevData.Cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 6, 0); }
                //人工检测数据
                if (data.Cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.Cell56.Value); }
                if (data.Cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.Cell57.Value); }
                if (data.Cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.Cell58.Value); }
                if (data.Cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.Cell59.Value); }
                if (data.Cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.Cell60.Value); }

                if (data.Cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.Cell61.Value); }
                if (data.Cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.Cell62.Value); }
                if (data.Cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.Cell63.Value); }
                if (data.Cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.Cell64.Value); }
                if (data.Cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.Cell65.Value); }
                if (data.Cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.Cell66.Value); }
                if (data.Cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.Cell67.Value); }
                if (data.Cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.Cell68.Value); }
                if (data.Cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.Cell69.Value); }
                if (data.Cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.Cell70.Value); }
                if (data.Cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.Cell71.Value); }
                if (data.Cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.Cell72.Value); }
                if (data.Cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.Cell73.Value); }
                if (data.Cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.Cell74.Value); }
                if (data.Cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.Cell75.Value); }
                if (data.Cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.Cell76.Value); }
                if (data.Cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.Cell77.Value); }
                if (data.Cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.Cell78.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell79 != null && prevData != null && prevData.Cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell79);
                        float prevVal = Convert.ToSingle(prevData.Cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.Cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.Cell80.Value); }
                if (data.Cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.Cell81.Value); }
                //人工检测数据
                if (data.Cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.Cell82.Value); }
                if (data.Cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.Cell83.Value); }
                if (data.Cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.Cell84.Value); }
                if (data.Cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.Cell85.Value); }
                if (data.Cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.Cell86.Value); }
                if (data.Cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.Cell87.Value); }

                if (data.Cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.Cell88.Value); }
                if (data.Cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.Cell89.Value); }
                if (data.Cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.Cell90.Value); }
                if (data.Cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.Cell91.Value); }
                if (data.Cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.Cell92.Value); }
                //if (data.Cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.Cell93.Value); }
                //if (data.Cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.Cell94.Value); }
                //if (data.Cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.Cell95.Value); }
                //if (data.Cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.Cell96.Value); }
                //if (data.Cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.Cell97.Value); }
                //if (data.Cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.Cell98.Value); }
                //if (data.Cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.Cell99.Value); }
                //if (data.Cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.Cell100.Value); }

                //Rang3
                if (data.Cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.Cell101.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell102 != null && prevData != null && prevData.Cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell102);
                        float prevVal = Convert.ToSingle(prevData.Cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.Cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.Cell103.Value); }
                if (data.Cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.Cell104.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell105 != null && prevData != null && prevData.Cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell105);
                        float prevVal = Convert.ToSingle(prevData.Cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.Cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.Cell106.Value); }
                if (data.Cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.Cell107.Value); }
                if (data.Cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.Cell108.Value); }
                if (data.Cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.Cell109.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell110 != null && prevData != null && prevData.Cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell110);
                        float prevVal = Convert.ToSingle(prevData.Cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.Cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.Cell111.Value); }
                if (data.Cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.Cell112.Value); }
                if (data.Cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.Cell113.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell114 != null && prevData != null && prevData.Cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell114);
                        float prevVal = Convert.ToSingle(prevData.Cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.Cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.Cell115.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell116 != null && prevData != null && prevData.Cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell116);
                        float prevVal = Convert.ToSingle(prevData.Cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.Cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.Cell117.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell118 != null && prevData != null && prevData.Cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell118);
                        float prevVal = Convert.ToSingle(prevData.Cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.Cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.Cell119.Value); }
                if (data.Cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.Cell120.Value); }
                if (data.Cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.Cell121.Value); }
                if (data.Cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.Cell122.Value); }
                if (data.Cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.Cell123.Value); }
                if (data.Cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.Cell124.Value); }
                if (data.Cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.Cell125.Value); }
                if (data.Cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.Cell126.Value); }
                if (data.Cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.Cell127.Value); }
                if (data.Cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.Cell128.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell129 != null && prevData != null && prevData.Cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell129);
                        float prevVal = Convert.ToSingle(prevData.Cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.Cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.Cell130.Value); }
                if (data.Cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.Cell131.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell132 != null && prevData != null && prevData.Cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell132);
                        float prevVal = Convert.ToSingle(prevData.Cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.Cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.Cell133.Value); }
                if (data.Cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.Cell134.Value); }
                //人工检测数据
                if (data.Cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.Cell135.Value); }
                if (data.Cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.Cell136.Value); }
                if (data.Cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.Cell137.Value); }
                if (data.Cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.Cell138.Value); }
                if (data.Cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.Cell139.Value); }
                if (data.Cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.Cell140.Value); }
                if (data.Cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.Cell141.Value); }

                //if (data.Cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.Cell142.Value); }
                //if (data.Cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.Cell143.Value); }
                //if (data.Cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.Cell144.Value); }
                //if (data.Cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.Cell145.Value); }
                //if (data.Cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.Cell146.Value); }
                //if (data.Cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.Cell147.Value); }
                //if (data.Cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.Cell148.Value); }
                //if (data.Cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.Cell149.Value); }
                //if (data.Cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.Cell150.Value); }

            }
            return true;
        }
        /// <summary>
        /// 写Xlsx数据  周
        /// </summary>
        private static bool WriteXlsxWeekly1(XSSFWorkbook srcWorkbook, CalculatedData?[] dataList, DateTime ReportDataTime)
        {
            ISheet srcSheet = srcWorkbook.GetSheetAt(1); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy-MM-dd");
            SetXlsxCellString(srcSheet, 0, 0, Temp);//记录日期
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 0; i < 8; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过
                int rowIndex = 5 + i;


                //模板还未提供
            }
            return true;
        }
        /// <summary>
        /// 写Xlsx数据  上月
        /// </summary>
        private static bool WriteXlsxMonthly(XSSFWorkbook srcWorkbook, CalculatedData?[] dataList, DateTime ReportDataTime)
        {
            ISheet srcSheet = srcWorkbook.GetSheetAt(0); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy-MM");
            SetXlsxCellString(srcSheet, 0, 0, Temp);//记录日期
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 0; i < 32; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过
                int rowIndex = 5 + i;


                //模板还未提供
            }
            return true;
        }
        /// <summary>
        /// 写Xlsx数据  去年年
        /// </summary>
        private static bool WriteXlsxYearly(XSSFWorkbook srcWorkbook, CalculatedData?[] dataList, DateTime ReportDataTime)
        {
            ISheet srcSheet = srcWorkbook.GetSheetAt(0); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy");
            SetXlsxCellString(srcSheet, 0, 0, Temp);//记录日期
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 0; i < 13; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过
                int rowIndex = 5 + i;

                //模板还未提供
            }
            return true;
        }
        /// <summary>
        /// 复制XLSX工作表（仅针对.xlsx，保留样式、合并单元格、列宽）
        /// </summary>
        private static ISheet CopyXlsxSheet3(XSSFWorkbook srcWorkbook, XSSFWorkbook destWorkbook, ISheet srcSheet, string newSheetName)
        {
            ISheet destSheet = destWorkbook.CreateSheet(newSheetName);

            // 1. 复制列宽（先获取模板表的最大列数，避免用srcSheet.LastCellNum）
            int maxColumnCount = 0;
            for (int rowIdx = 0; rowIdx <= srcSheet.LastRowNum; rowIdx++)
            {
                IRow srcRow = srcSheet.GetRow(rowIdx);
                if (srcRow != null && srcRow.LastCellNum > maxColumnCount)
                {
                    maxColumnCount = srcRow.LastCellNum; // 用IRow的LastCellNum
                }
            }
            for (int col = 0; col < maxColumnCount; col++)
            {
                destSheet.SetColumnWidth(col, srcSheet.GetColumnWidth(col));
            }

            // 2. 复制行和单元格（修复日期赋值错误）
            for (int rowIdx = 0; rowIdx <= srcSheet.LastRowNum; rowIdx++)
            {
                IRow srcRow = srcSheet.GetRow(rowIdx);
                IRow destRow = destSheet.CreateRow(rowIdx);

                if (srcRow != null)
                {
                    destRow.Height = srcRow.Height;

                    // 复制单元格（遍历行的LastCellNum）
                    for (int CellIdx = 0; CellIdx < srcRow.LastCellNum; CellIdx++)
                    {
                        ICell srcCell = srcRow.GetCell(CellIdx);
                        if (srcCell != null)
                        {
                            ICell destCell = destRow.CreateCell(CellIdx);
                            destCell.CellStyle = srcCell.CellStyle;
                            CopyCellValue(srcCell, destCell); // 修复日期赋值逻辑
                        }
                    }
                }
            }

            // 3. 复制合并单元格
            foreach (CellRangeAddress region in srcSheet.MergedRegions)
            {
                destSheet.AddMergedRegion(region);
            }

            return destSheet;
        }
        /// <summary>
        /// 复制单元格值（适配不同数据类型）
        /// </summary>
        private static void CopyCellValue(ICell srcCell, ICell destCell)
        {
            switch (srcCell.CellType)
            {
                case CellType.String:
                    destCell.SetCellValue(srcCell.StringCellValue);
                    break;
                case CellType.Numeric:
                    // 重点修复：显式处理日期类型，避免调用SetCellValue(double)
                    if (DateUtil.IsCellDateFormatted(srcCell))
                    {
                        // 处理可空DateTime：判断HasValue后取Value
                        DateTime? dateValue = srcCell.DateCellValue;
                        destCell.SetCellValue(dateValue.HasValue ? dateValue.Value : DateTime.MinValue);
                    }
                    else
                    {
                        destCell.SetCellValue(srcCell.NumericCellValue);
                    }
                    break;
                case CellType.Boolean:
                    destCell.SetCellValue(srcCell.BooleanCellValue);
                    break;
                case CellType.Formula:
                    destCell.SetCellFormula(srcCell.CellFormula);
                    break;
                default:
                    destCell.SetCellValue(srcCell.ToString());
                    break;
            }
        }

        /// <summary>
        /// 给XLSX单元格赋值（封装逻辑，简化调用）
        /// </summary>
        private static void SetXlsxCellValue(ISheet sheet, int rowIdx, int colIdx, float value)
        {
            // 获取或创建行
            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
            // 获取或创建单元格
            ICell Cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            // 赋值
            Cell.SetCellValue(value);

        }
        //写日期
        private static void SetXlsxCellString(ISheet sheet, int rowIdx, int colIdx, string value)
        {
            // 获取或创建行
            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
            // 获取或创建单元格
            ICell Cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            // 赋值
            Cell.SetCellValue(value);

        }

        public async Task<List<SourceData>> GetSourceData(DateTime StartTime, DateTime EndtTime)
        {
            var result = await _sourceData.GetByDateTimeRangeAsync(StartTime, EndtTime);
            return result;
        }

        public async Task<bool> UpdateSourceDataFieldAsync(string dateStr, int hour, string prop, string valueStr)
        {

            if (!DateTime.TryParse(dateStr, out DateTime targetDate))
            {
                throw new ArgumentException($"日期格式错误，要求yyyy-MM-dd，当前值：{dateStr}", nameof(dateStr));
            }

            // 构建目标时间（精确到小时，用于筛选记录）
            DateTime targetDateTime = new DateTime(targetDate.Year, targetDate.Month, targetDate.Day, hour, 0, 0);

            // 转换值为float?类型
            if (!float.TryParse(valueStr, out float value))
            {
                throw new ArgumentException($"值转换失败，要求浮点数字符串，当前值：{valueStr}", nameof(valueStr));
            }
            float? targetValue = value; // 兼容nullable float类型

            // 校验字段名是否存在
            PropertyInfo? propInfo = typeof(SourceData).GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
            if (propInfo == null)
            {
                throw new ArgumentException($"SourceData 不存在字段：{prop}", nameof(prop));
            }
            if (propInfo.PropertyType != typeof(float?))
            {
                throw new ArgumentException($"字段{prop}类型不是float?，不支持修改", nameof(prop));
            }
            // 方式1：精确匹配时间（可根据业务调整为时间范围）
            var targetData = await _sourceData.Db
                .FirstOrDefaultAsync(d => d.ReportedTime >= targetDateTime
                                        && d.ReportedTime < targetDateTime.AddHours(1));

            if (targetData == null)
            {
                throw new KeyNotFoundException($"未找到{targetDateTime:yyyy-MM-dd HH:mm}时间段的 SourceData 记录");
            }
            try
            {
                propInfo.SetValue(targetData, targetValue);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"设置字段{prop}值失败：{ex.Message}", ex);
            }
            await _sourceData.Update(targetData); // 标记实体为修改状态
            await _dbContext.SaveChangesAsync(); // 提交到数据库

            return true;
        }

    }


}
