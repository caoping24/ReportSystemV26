using CenterBackend.Dto;
using CenterBackend.IReportServices;
using CenterReport.Repository;
using CenterReport.Repository.Models;
using Masuit.Tools.DateTimeExt;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.X509;
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
                             IHttpContextAccessor httpContextAccessor,
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
            var reportedTime = startTime.Date;//记录是那一天的数据
            var target = _calculatedDatas.db.FirstOrDefault(r => r.Type == type && r.reportedTime == reportedTime);
            bool isNewRecord = (target == null);
            if (isNewRecord)
            {
                target = new CalculatedData
                {
                    Type = type,
                    reportedTime = reportedTime,
                    createdtime = DateTime.Now // 新增时初始化创建时间
                };
            }
            else
            {
                target.createdtime = DateTime.Now; // 更新时刷新创建时间（或改updateTime更合理）
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
                    var dailyDataList = await _sourceData.GetByDataTimeAsync(startTime, stopTime);
                    if (dailyDataList.Count == 0) return false;
                    await DayDataCalculate(target, dailyDataList);
                    break;

                case 2:
                    var weeklyDataList = await _calculatedDatas.GetByDataTimeAsync(startTime, stopTime, 1);
                    if (weeklyDataList.Count == 0) return false;
                    await WeekDataCalculate(target, weeklyDataList);
                    break;

                case 3:
                    var monthlyDataList = await _calculatedDatas.GetByDataTimeAsync(startTime, stopTime, 1);
                    if (monthlyDataList.Count == 0) return false;
                    await MonthDataCalculate(target, monthlyDataList);
                    break;

                case 4:
                    var yearlyDataList = await _calculatedDatas.GetByDataTimeAsync(startTime, stopTime, 3);
                    if (yearlyDataList.Count == 0) return false;
                    await YearDataCalculate(target, yearlyDataList);
                    break;

                default:
                    return false;
            }

            return true;
        }
        private async Task DayDataCalculate(CalculatedData target, List<SourceData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义
            //(dataList.Last().cell2 ?? 0) - (dataList.First().cell2 ?? 0);//差值
            // dataList.Select(x => x.cell3 ?? 0).Sum();//总和
            target.cell1 = dataList.Select(x => x.cell1 ?? 0).Average();//平均值
            target.cell2 = dataList.Select(x => x.cell2 ?? 0).Average();
            target.cell3 = dataList.Select(x => x.cell3 ?? 0).Average();
            target.cell4 = dataList.Last().cell4 - dataList.First().cell4;//差值
            target.cell5 = dataList.Last().cell5 - dataList.First().cell5;//差值
            target.cell6 = dataList.Select(x => x.cell6 ?? 0).Average();
            target.cell7 = dataList.Select(x => x.cell7 ?? 0).Average();
            target.cell8 = dataList.Last().cell8 - dataList.First().cell8;//差值
            target.cell9 = dataList.Select(x => x.cell9 ?? 0).Average();
            target.cell10 = dataList.Select(x => x.cell10 ?? 0).Average();
            target.cell11 = dataList.Select(x => x.cell11 ?? 0).Average();
            target.cell12 = dataList.Select(x => x.cell12 ?? 0).Average();
            target.cell13 = dataList.Select(x => x.cell13 ?? 0).Average();
            target.cell14 = dataList.Last().cell14 - dataList.First().cell14;//差值
            target.cell15 = dataList.Select(x => x.cell15 ?? 0).Average();
            target.cell16 = dataList.Last().cell16 - dataList.First().cell16;//差值
            target.cell17 = dataList.Select(x => x.cell17 ?? 0).Average();
            target.cell18 = dataList.Select(x => x.cell18 ?? 0).Average();
            target.cell19 = dataList.Select(x => x.cell19 ?? 0).Average();
            target.cell20 = dataList.Last().cell20 - dataList.First().cell20;//差值
            target.cell21 = dataList.Select(x => x.cell21 ?? 0).Average();
            target.cell22 = dataList.Select(x => x.cell22 ?? 0).Average();
            target.cell23 = dataList.Select(x => x.cell23 ?? 0).Average();
            target.cell24 = dataList.Last().cell24;//最后一个值
            target.cell25 = dataList.Select(x => x.cell25 ?? 0).Average();
            target.cell26 = dataList.Select(x => x.cell26 ?? 0).Average();
            target.cell27 = dataList.Select(x => x.cell27 ?? 0).Average();
            target.cell28 = dataList.Select(x => x.cell28 ?? 0).Average();
            //target.cell29 = dataList.Select(x => x.cell29 ?? 0).Average();//人工填写
            //target.cell30 = dataList.Select(x => x.cell30 ?? 0).Average();//人工填写
            //target.cell31 = dataList.Select(x => x.cell31 ?? 0).Average();//人工填写
            //target.cell32 = dataList.Select(x => x.cell32 ?? 0).Average();//人工填写
            //target.cell33 = dataList.Select(x => x.cell33 ?? 0).Average();//人工填写
            //target.cell34 = dataList.Select(x => x.cell34 ?? 0).Average();//人工填写
            //target.cell35 = dataList.Select(x => x.cell35 ?? 0).Average();//人工填写
            target.cell36 = dataList.Select(x => x.cell36 ?? 0).Average();
            target.cell37 = dataList.Last().cell37 - dataList.First().cell37;//差值
            target.cell38 = dataList.Select(x => x.cell38 ?? 0).Average();
            target.cell39 = dataList.Select(x => x.cell39 ?? 0).Average();
            target.cell40 = dataList.Select(x => x.cell40 ?? 0).Average();
            target.cell41 = dataList.Select(x => x.cell41 ?? 0).Average();
            target.cell42 = dataList.Last().cell42 - dataList.First().cell42;//差值
            //target.cell43 = dataList.Select(x => x.cell43 ?? 0).Average();
            //target.cell44 = dataList.Select(x => x.cell44 ?? 0).Average();
            //target.cell45 = dataList.Select(x => x.cell45 ?? 0).Average();
            //target.cell46 = dataList.Select(x => x.cell46 ?? 0).Average();
            //target.cell47 = dataList.Select(x => x.cell47 ?? 0).Average();
            //target.cell48 = dataList.Select(x => x.cell48 ?? 0).Average();
            //target.cell49 = dataList.Select(x => x.cell49 ?? 0).Average();
            //target.cell50 = dataList.Select(x => x.cell50 ?? 0).Average();
            // 第二组：cell51-cell100
            target.cell51 = dataList.Select(x => x.cell51 ?? 0).Average();
            target.cell52 = dataList.Select(x => x.cell52 ?? 0).Average();
            target.cell53 = dataList.Select(x => x.cell53 ?? 0).Average();
            target.cell54 = dataList.Select(x => x.cell54 ?? 0).Average();
            target.cell55 = dataList.Last().cell55 - dataList.First().cell55;//差值
            //target.cell56 = dataList.Select(x => x.cell56 ?? 0).Average();
            //target.cell57 = dataList.Select(x => x.cell57 ?? 0).Average();
            //target.cell58 = dataList.Select(x => x.cell58 ?? 0).Average();
            //target.cell59 = dataList.Select(x => x.cell59 ?? 0).Average();
            //target.cell60 = dataList.Select(x => x.cell60 ?? 0).Average();
            target.cell61 = dataList.Select(x => x.cell61 ?? 0).Average();
            target.cell62 = dataList.Select(x => x.cell62 ?? 0).Average();
            target.cell63 = dataList.Select(x => x.cell63 ?? 0).Average();
            target.cell64 = dataList.Select(x => x.cell64 ?? 0).Average();
            target.cell65 = dataList.Select(x => x.cell65 ?? 0).Average();
            target.cell66 = dataList.Select(x => x.cell66 ?? 0).Average();
            target.cell67 = dataList.Select(x => x.cell67 ?? 0).Average();
            target.cell68 = dataList.Select(x => x.cell68 ?? 0).Average();
            target.cell69 = dataList.Select(x => x.cell69 ?? 0).Average();
            target.cell70 = dataList.Select(x => x.cell70 ?? 0).Average();
            target.cell71 = dataList.Select(x => x.cell71 ?? 0).Average();
            target.cell72 = dataList.Select(x => x.cell72 ?? 0).Average();
            target.cell73 = dataList.Select(x => x.cell73 ?? 0).Average();
            target.cell74 = dataList.Select(x => x.cell74 ?? 0).Average();
            target.cell75 = dataList.Select(x => x.cell75 ?? 0).Average();
            target.cell76 = dataList.Select(x => x.cell76 ?? 0).Average();
            target.cell77 = dataList.Select(x => x.cell77 ?? 0).Average();
            target.cell78 = dataList.Select(x => x.cell78 ?? 0).Average();
            target.cell79 = dataList.Last().cell79 - dataList.First().cell79;//差值
            target.cell80 = dataList.Select(x => x.cell80 ?? 0).Average();
            target.cell81 = dataList.Select(x => x.cell81 ?? 0).Average();
            //target.cell82 = dataList.Select(x => x.cell82 ?? 0).Average();
            //target.cell83 = dataList.Select(x => x.cell83 ?? 0).Average();
            //target.cell84 = dataList.Select(x => x.cell84 ?? 0).Average();
            //target.cell85 = dataList.Select(x => x.cell85 ?? 0).Average();
            //target.cell86 = dataList.Select(x => x.cell86 ?? 0).Average();
            //target.cell87 = dataList.Select(x => x.cell87 ?? 0).Average();
            target.cell88 = dataList.Select(x => x.cell88 ?? 0).Average();
            target.cell89 = dataList.Select(x => x.cell89 ?? 0).Average();
            target.cell90 = dataList.Select(x => x.cell90 ?? 0).Average();
            target.cell91 = dataList.Select(x => x.cell91 ?? 0).Average();
            target.cell92 = dataList.Select(x => x.cell92 ?? 0).Average();
            //target.cell93 = dataList.Select(x => x.cell93 ?? 0).Average();
            //target.cell94 = dataList.Select(x => x.cell94 ?? 0).Average();
            //target.cell95 = dataList.Select(x => x.cell95 ?? 0).Average();
            //target.cell96 = dataList.Select(x => x.cell96 ?? 0).Average();
            //target.cell97 = dataList.Select(x => x.cell97 ?? 0).Average();
            //target.cell98 = dataList.Select(x => x.cell98 ?? 0).Average();
            //target.cell99 = dataList.Select(x => x.cell99 ?? 0).Average();
            //target.cell100 = dataList.Select(x => x.cell100 ?? 0).Average();
            // 第三组：cell101-cell150
            target.cell101 = dataList.Select(x => x.cell101 ?? 0).Average();
            target.cell102 = dataList.Last().cell102 - dataList.First().cell102;//差值
            target.cell103 = dataList.Select(x => x.cell103 ?? 0).Average();
            target.cell104 = dataList.Last().cell104 - dataList.First().cell104;//差值
            target.cell105 = dataList.Select(x => x.cell105 ?? 0).Average();
            target.cell106 = dataList.Select(x => x.cell106 ?? 0).Average();
            target.cell107 = dataList.Select(x => x.cell107 ?? 0).Average();
            target.cell108 = dataList.Select(x => x.cell108 ?? 0).Average();
            target.cell109 = dataList.Select(x => x.cell109 ?? 0).Average();
            target.cell110 = dataList.Last().cell110 - dataList.First().cell110;//差值
            target.cell111 = dataList.Select(x => x.cell111 ?? 0).Average();
            target.cell112 = dataList.Select(x => x.cell112 ?? 0).Average();
            target.cell113 = dataList.Select(x => x.cell113 ?? 0).Average();
            target.cell114 = dataList.Last().cell114 - dataList.First().cell114;//差值
            target.cell115 = dataList.Select(x => x.cell115 ?? 0).Average();
            target.cell116 = dataList.Last().cell116 - dataList.First().cell116;//差值
            target.cell117 = dataList.Select(x => x.cell117 ?? 0).Average();
            target.cell118 = dataList.Last().cell118 - dataList.First().cell118;//差值
            target.cell119 = dataList.Select(x => x.cell119 ?? 0).Average();
            target.cell120 = dataList.Select(x => x.cell120 ?? 0).Average();
            target.cell121 = dataList.Select(x => x.cell121 ?? 0).Average();
            target.cell122 = dataList.Select(x => x.cell122 ?? 0).Average();
            target.cell123 = dataList.Select(x => x.cell123 ?? 0).Average();
            target.cell124 = dataList.Select(x => x.cell124 ?? 0).Average();
            target.cell125 = dataList.Select(x => x.cell125 ?? 0).Average();
            target.cell126 = dataList.Select(x => x.cell126 ?? 0).Average();
            target.cell127 = dataList.Select(x => x.cell127 ?? 0).Average();
            target.cell128 = dataList.Select(x => x.cell128 ?? 0).Average();
            //target.cell129 = dataList.Last().cell129 - dataList.First().cell129;//差值
            target.cell130 = dataList.Select(x => x.cell130 ?? 0).Average();
            target.cell131 = dataList.Select(x => x.cell131 ?? 0).Average();
            target.cell132 = dataList.Last().cell132 - dataList.First().cell132;//差值
            target.cell133 = dataList.Select(x => x.cell133 ?? 0).Average();
            target.cell134 = dataList.Select(x => x.cell134 ?? 0).Average();
            //target.cell135 = dataList.Select(x => x.cell135 ?? 0).Average();
            //target.cell136 = dataList.Select(x => x.cell136 ?? 0).Average();
            //target.cell137 = dataList.Select(x => x.cell137 ?? 0).Average();
            //target.cell138 = dataList.Select(x => x.cell138 ?? 0).Average();
            //target.cell139 = dataList.Select(x => x.cell139 ?? 0).Average();
            //target.cell140 = dataList.Select(x => x.cell140 ?? 0).Average();
            //target.cell141 = dataList.Select(x => x.cell141 ?? 0).Average();
            //target.cell142 = dataList.Select(x => x.cell142 ?? 0).Average();
            //target.cell143 = dataList.Select(x => x.cell143 ?? 0).Average();
            //target.cell144 = dataList.Select(x => x.cell144 ?? 0).Average();
            //target.cell145 = dataList.Select(x => x.cell145 ?? 0).Average();
            //target.cell146 = dataList.Select(x => x.cell146 ?? 0).Average();
            //target.cell147 = dataList.Select(x => x.cell147 ?? 0).Average();
            //target.cell148 = dataList.Select(x => x.cell148 ?? 0).Average();
            //target.cell149 = dataList.Select(x => x.cell149 ?? 0).Average();
            //target.cell150 = dataList.Select(x => x.cell150 ?? 0).Average();

        }
        private async Task WeekDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.cell1 = dataList.Select(x => x.cell1 ?? 0).Average();//平均值
        }
        private async Task MonthDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.cell1 = dataList.Select(x => x.cell1 ?? 0).Average();//平均值
        }
        private async Task YearDataCalculate(CalculatedData target, List<CalculatedData> dataList)
        {

            target.PH = 80;//暂时没有特殊意义

            target.cell1 = dataList.Select(x => x.cell1 ?? 0).Average();//平均值
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
                        dataList = await _sourceData.GetByDayAsync(ReportTime);
                        if (dataList == null || !dataList.Any())
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
                        int daysToThisMonday = ((int)currentDayOfWeek.DayOfWeek + 6) % 7 ;
                        StartTime = currentDayOfWeek.AddDays(-daysToThisMonday);
                        StopTime = StartTime.AddDays(6).AddHours(23).AddMinutes(59);

                        dataList2 = await _calculatedDatas.GetByDataTimeAsync(StartTime, StopTime, 2);
                        if (dataList2 == null || !dataList2.Any())
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
                        dataList2 = await _calculatedDatas.GetByDataTimeAsync(StartTime, StopTime, 3);
                        if (dataList2 == null || !dataList2.Any())
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
                        dataList2 = await _calculatedDatas.GetByDataTimeAsync(StartTime, StopTime, 4);
                        if (dataList2 == null || !dataList2.Any())
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
                var existingRecord = _reportRecord.db.FirstOrDefault(r => r.Type == Type && r.reportedTime == TempRepoetName);
                if (existingRecord == null)//无记录则新增，有记录则更新
                {
                    var temp = new ReportRecord();
                    temp.Type = Type;
                    temp.reportedTime = TempRepoetName;
                    temp.createdtime = DateTime.Now;
                    await _reportRecord.AddAsync(temp);
                }
                else
                {
                    existingRecord.createdtime = DateTime.Now;
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
            if (sourceDataList == null || !sourceDataList.Any()) return targetArray;
            startTime = startTime.Date.AddHours(8);
            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.createdtime - startTime).TotalHours;
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
            if (sourceDataList == null || !sourceDataList.Any()) return targetArray;

            startTime = startTime.AddHours(8); //上周一8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.createdtime - startTime).TotalDays;
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
            if (sourceDataList == null || !sourceDataList.Any()) return targetArray;

            startTime = startTime.AddHours(8); //上月1号8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;
                double hourDiff = (data.createdtime - startTime).TotalDays;
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
            if (sourceDataList == null || !sourceDataList.Any()) return targetArray;

            startTime = startTime.AddHours(8); //去年1月1号8点

            foreach (var data in sourceDataList)
            {
                if (data == null) continue;

                int matchIndex = (data.createdtime.Year - startTime.Year) * 12 + (data.createdtime.Month - startTime.Month);
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
                if (data.cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.cell1.Value); }
                if (data.cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.cell2.Value); }
                if (data.cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.cell3.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell4 != null && prevData != null && prevData != null && prevData.cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell4);
                        float prevVal = Convert.ToSingle(prevData.cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell5 != null && prevData != null && prevData.cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell5);
                        float prevVal = Convert.ToSingle(prevData.cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.cell6.Value); }
                if (data.cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.cell7.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell8 != null && prevData != null && prevData.cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell8);
                        float prevVal = Convert.ToSingle(prevData.cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.cell9.Value); }
                if (data.cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.cell10.Value); }
                if (data.cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.cell11.Value); }
                if (data.cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.cell12.Value); }
                if (data.cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.cell13.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell14 != null && prevData != null && prevData.cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell14);
                        float prevVal = Convert.ToSingle(prevData.cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.cell15.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell16 != null && prevData != null && prevData.cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell16);
                        float prevVal = Convert.ToSingle(prevData.cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.cell17.Value); }
                if (data.cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.cell18.Value); }
                if (data.cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.cell19.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell20 != null && prevData != null && prevData.cell20 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell20);
                        float prevVal = Convert.ToSingle(prevData.cell20);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 21, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.cell21.Value); }
                if (data.cell22 != null) //摩尔比 大于0 小于2 
                { 
                    if (data.cell22.Value < 2 )
                    SetXlsxCellValue(srcSheet, Range1, 23, data.cell22.Value); 
                }
                if (i == 12)
                {
                    if (data.cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.cell23.Value); }//只记录最后一个值
                }
                if (data.cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.cell24.Value); }

                if (data.cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.cell25.Value); }
                if (data.cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.cell26.Value); }
                if (data.cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.cell27.Value); }
                if (data.cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.cell28.Value); }
                //人工检测数据
                if (data.cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.cell29.Value); }
                if (data.cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.cell30.Value); }
                if (data.cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.cell31.Value); }
                if (data.cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.cell32.Value); }
                if (data.cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.cell33.Value); }
                if (data.cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.cell34.Value); }
                if (data.cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.cell35.Value); }

                if (data.cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.cell36.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell37 != null && prevData != null && prevData.cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell37);
                        float prevVal = Convert.ToSingle(prevData.cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.cell38.Value); }
                if (data.cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.cell39.Value); }
                if (data.cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.cell40.Value); }
                if (data.cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.cell41.Value * 1000 ); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell42 != null && prevData != null && prevData.cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell42);
                        float prevVal = Convert.ToSingle(prevData.cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.cell43.Value); }
                //if (data.cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.cell44.Value); }
                //if (data.cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.cell45.Value); }
                //if (data.cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.cell46.Value); }
                //if (data.cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.cell47.Value); }
                //if (data.cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.cell48.Value); }
                //if (data.cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.cell49.Value); }
                //if (data.cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.cell50.Value); }

                //Rang2
                if (data.cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.cell51.Value); }
                if (data.cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.cell52.Value); }
                if (data.cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.cell53.Value); }
                if (data.cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.cell54.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell55 != null && prevData != null && prevData.cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell55);
                        float prevVal = Convert.ToSingle(prevData.cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                //人工检测数据
                if (data.cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.cell56.Value); }
                if (data.cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.cell57.Value); }
                if (data.cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.cell58.Value); }
                if (data.cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.cell59.Value); }
                if (data.cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.cell60.Value); }

                if (data.cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.cell61.Value); }
                if (data.cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.cell62.Value); }
                if (data.cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.cell63.Value); }
                if (data.cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.cell64.Value); }
                if (data.cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.cell65.Value); }
                if (data.cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.cell66.Value); }
                if (data.cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.cell67.Value); }
                if (data.cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.cell68.Value); }
                if (data.cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.cell69.Value); }
                if (data.cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.cell70.Value); }
                if (data.cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.cell71.Value); }
                if (data.cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.cell72.Value); }
                if (data.cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.cell73.Value); }
                if (data.cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.cell74.Value); }
                if (data.cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.cell75.Value); }
                if (data.cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.cell76.Value); }
                if (data.cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.cell77.Value); }
                if (data.cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.cell78.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell79 != null && prevData != null && prevData.cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell79);
                        float prevVal = Convert.ToSingle(prevData.cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.cell80.Value); }
                if (data.cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.cell81.Value); }
                //人工检测数据
                if (data.cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.cell82.Value); }
                if (data.cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.cell83.Value); }
                if (data.cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.cell84.Value); }
                if (data.cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.cell85.Value); }
                if (data.cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.cell86.Value); }
                if (data.cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.cell87.Value); }

                if (data.cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.cell88.Value); }
                if (data.cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.cell89.Value); }
                if (data.cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.cell90.Value); }
                if (data.cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.cell91.Value); }
                if (data.cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.cell92.Value); }
                //if (data.cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.cell93.Value); }
                //if (data.cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.cell94.Value); }
                //if (data.cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.cell95.Value); }
                //if (data.cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.cell96.Value); }
                //if (data.cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.cell97.Value); }
                //if (data.cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.cell98.Value); }
                //if (data.cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.cell99.Value); }
                //if (data.cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.cell100.Value); }

                //Rang3
                if (data.cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.cell101.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell102 != null && prevData != null && prevData.cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell102);
                        float prevVal = Convert.ToSingle(prevData.cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.cell103.Value); }
                if (data.cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.cell104.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell105 != null && prevData != null && prevData.cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell105);
                        float prevVal = Convert.ToSingle(prevData.cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.cell106.Value); }
                if (data.cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.cell107.Value); }
                if (data.cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.cell108.Value); }
                if (data.cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.cell109.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell110 != null && prevData != null && prevData.cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell110);
                        float prevVal = Convert.ToSingle(prevData.cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.cell111.Value); }
                if (data.cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.cell112.Value); }
                if (data.cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.cell113.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell114 != null && prevData != null && prevData.cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell114);
                        float prevVal = Convert.ToSingle(prevData.cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.cell115.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell116 != null && prevData != null && prevData.cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell116);
                        float prevVal = Convert.ToSingle(prevData.cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.cell117.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell118 != null && prevData != null && prevData.cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell118);
                        float prevVal = Convert.ToSingle(prevData.cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.cell119.Value); }
                if (data.cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.cell120.Value); }
                if (data.cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.cell121.Value); }
                if (data.cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.cell122.Value); }
                if (data.cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.cell123.Value); }
                if (data.cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.cell124.Value); }
                if (data.cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.cell125.Value); }
                if (data.cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.cell126.Value); }
                if (data.cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.cell127.Value); }
                if (data.cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.cell128.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell129 != null && prevData != null && prevData.cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell129);
                        float prevVal = Convert.ToSingle(prevData.cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.cell130.Value); }
                if (data.cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.cell131.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell132 != null && prevData != null && prevData.cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell132);
                        float prevVal = Convert.ToSingle(prevData.cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.cell133.Value); }
                if (data.cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.cell134.Value); }
                //人工检测数据
                if (data.cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.cell135.Value); }
                if (data.cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.cell136.Value); }
                if (data.cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.cell137.Value); }
                if (data.cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.cell138.Value); }
                if (data.cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.cell139.Value); }
                if (data.cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.cell140.Value); }
                if (data.cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.cell141.Value); }

                //if (data.cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.cell142.Value); }
                //if (data.cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.cell143.Value); }
                //if (data.cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.cell144.Value); }
                //if (data.cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.cell145.Value); }
                //if (data.cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.cell146.Value); }
                //if (data.cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.cell147.Value); }
                //if (data.cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.cell148.Value); }
                //if (data.cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.cell149.Value); }
                //if (data.cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.cell150.Value); }

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
                if (data.cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.cell1.Value); }
                if (data.cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.cell2.Value); }
                if (data.cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.cell3.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell4 != null && prevData != null && prevData.cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell4);
                        float prevVal = Convert.ToSingle(prevData.cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell5 != null && prevData != null && prevData.cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell5);
                        float prevVal = Convert.ToSingle(prevData.cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.cell6.Value); }
                if (data.cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.cell7.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell8 != null && prevData != null && prevData.cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell8);
                        float prevVal = Convert.ToSingle(prevData.cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.cell9.Value); }
                if (data.cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.cell10.Value); }
                if (data.cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.cell11.Value); }
                if (data.cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.cell12.Value); }
                if (data.cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.cell13.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell14 != null && prevData != null && prevData.cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell14);
                        float prevVal = Convert.ToSingle(prevData.cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.cell15.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell16 != null && prevData != null && prevData.cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell16);
                        float prevVal = Convert.ToSingle(prevData.cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.cell17.Value); }
                if (data.cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.cell18.Value); }
                if (data.cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.cell19.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                        var prevData = dataList.ElementAt(i - 1);
                        if (data.cell20 != null && prevData != null && prevData.cell20 != null)
                        {
                            float currentVal = Convert.ToSingle(data.cell20);
                            float prevVal = Convert.ToSingle(prevData.cell20);
                            float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                            SetXlsxCellValue(srcSheet, Range1, 21, result);
                        }
                    }
                    else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.cell21.Value); }
                if (data.cell22 != null) //摩尔比 大于0 小于2 
                {
                    if (data.cell22.Value < 2)
                        SetXlsxCellValue(srcSheet, Range1, 23, data.cell22.Value);
                }
                if (i == 24)
                {
                    if (data.cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.cell23.Value); }//只记录最后一个值
                }
                if (data.cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.cell24.Value); }

                if (data.cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.cell25.Value); }
                if (data.cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.cell26.Value); }
                if (data.cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.cell27.Value); }
                if (data.cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.cell28.Value); }
                //人工检测数据
                if (data.cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.cell29.Value); }
                if (data.cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.cell30.Value); }
                if (data.cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.cell31.Value); }
                if (data.cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.cell32.Value); }
                if (data.cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.cell33.Value); }
                if (data.cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.cell34.Value); }
                if (data.cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.cell35.Value); }

                if (data.cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.cell36.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell37 != null && prevData != null && prevData.cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell37);
                        float prevVal = Convert.ToSingle(prevData.cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.cell38.Value); }
                if (data.cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.cell39.Value); }
                if (data.cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.cell40.Value); }
                if (data.cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.cell41.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell42 != null && prevData != null && prevData.cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell42);
                        float prevVal = Convert.ToSingle(prevData.cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.cell43.Value); }
                //if (data.cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.cell44.Value); }
                //if (data.cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.cell45.Value); }
                //if (data.cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.cell46.Value); }
                //if (data.cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.cell47.Value); }
                //if (data.cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.cell48.Value); }
                //if (data.cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.cell49.Value); }
                //if (data.cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.cell50.Value); }

                //Rang2
                if (data.cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.cell51.Value); }
                if (data.cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.cell52.Value); }
                if (data.cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.cell53.Value); }
                if (data.cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.cell54.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell55 != null && prevData != null && prevData.cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell55);
                        float prevVal = Convert.ToSingle(prevData.cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 6, 0); }
                //人工检测数据
                if (data.cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.cell56.Value); }
                if (data.cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.cell57.Value); }
                if (data.cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.cell58.Value); }
                if (data.cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.cell59.Value); }
                if (data.cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.cell60.Value); }

                if (data.cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.cell61.Value); }
                if (data.cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.cell62.Value); }
                if (data.cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.cell63.Value); }
                if (data.cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.cell64.Value); }
                if (data.cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.cell65.Value); }
                if (data.cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.cell66.Value); }
                if (data.cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.cell67.Value); }
                if (data.cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.cell68.Value); }
                if (data.cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.cell69.Value); }
                if (data.cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.cell70.Value); }
                if (data.cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.cell71.Value); }
                if (data.cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.cell72.Value); }
                if (data.cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.cell73.Value); }
                if (data.cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.cell74.Value); }
                if (data.cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.cell75.Value); }
                if (data.cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.cell76.Value); }
                if (data.cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.cell77.Value); }
                if (data.cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.cell78.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell79 != null && prevData != null && prevData.cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell79);
                        float prevVal = Convert.ToSingle(prevData.cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.cell80.Value); }
                if (data.cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.cell81.Value); }
                //人工检测数据
                if (data.cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.cell82.Value); }
                if (data.cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.cell83.Value); }
                if (data.cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.cell84.Value); }
                if (data.cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.cell85.Value); }
                if (data.cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.cell86.Value); }
                if (data.cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.cell87.Value); }

                if (data.cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.cell88.Value); }
                if (data.cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.cell89.Value); }
                if (data.cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.cell90.Value); }
                if (data.cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.cell91.Value); }
                if (data.cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.cell92.Value); }
                //if (data.cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.cell93.Value); }
                //if (data.cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.cell94.Value); }
                //if (data.cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.cell95.Value); }
                //if (data.cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.cell96.Value); }
                //if (data.cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.cell97.Value); }
                //if (data.cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.cell98.Value); }
                //if (data.cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.cell99.Value); }
                //if (data.cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.cell100.Value); }

                //Rang3
                if (data.cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.cell101.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell102 != null && prevData != null && prevData.cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell102);
                        float prevVal = Convert.ToSingle(prevData.cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.cell103.Value); }
                if (data.cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.cell104.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell105 != null && prevData != null && prevData.cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell105);
                        float prevVal = Convert.ToSingle(prevData.cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.cell106.Value); }
                if (data.cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.cell107.Value); }
                if (data.cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.cell108.Value); }
                if (data.cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.cell109.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell110 != null && prevData != null && prevData.cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell110);
                        float prevVal = Convert.ToSingle(prevData.cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.cell111.Value); }
                if (data.cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.cell112.Value); }
                if (data.cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.cell113.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell114 != null && prevData != null && prevData.cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell114);
                        float prevVal = Convert.ToSingle(prevData.cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.cell115.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell116 != null && prevData != null && prevData.cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell116);
                        float prevVal = Convert.ToSingle(prevData.cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.cell117.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell118 != null && prevData != null && prevData.cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell118);
                        float prevVal = Convert.ToSingle(prevData.cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.cell119.Value); }
                if (data.cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.cell120.Value); }
                if (data.cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.cell121.Value); }
                if (data.cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.cell122.Value); }
                if (data.cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.cell123.Value); }
                if (data.cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.cell124.Value); }
                if (data.cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.cell125.Value); }
                if (data.cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.cell126.Value); }
                if (data.cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.cell127.Value); }
                if (data.cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.cell128.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell129 != null && prevData != null && prevData.cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell129);
                        float prevVal = Convert.ToSingle(prevData.cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.cell130.Value); }
                if (data.cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.cell131.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.cell132 != null && prevData != null && prevData.cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.cell132);
                        float prevVal = Convert.ToSingle(prevData.cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.cell133.Value); }
                if (data.cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.cell134.Value); }
                //人工检测数据
                if (data.cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.cell135.Value); }
                if (data.cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.cell136.Value); }
                if (data.cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.cell137.Value); }
                if (data.cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.cell138.Value); }
                if (data.cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.cell139.Value); }
                if (data.cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.cell140.Value); }
                if (data.cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.cell141.Value); }

                //if (data.cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.cell142.Value); }
                //if (data.cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.cell143.Value); }
                //if (data.cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.cell144.Value); }
                //if (data.cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.cell145.Value); }
                //if (data.cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.cell146.Value); }
                //if (data.cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.cell147.Value); }
                //if (data.cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.cell148.Value); }
                //if (data.cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.cell149.Value); }
                //if (data.cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.cell150.Value); }

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
                    for (int cellIdx = 0; cellIdx < srcRow.LastCellNum; cellIdx++)
                    {
                        ICell srcCell = srcRow.GetCell(cellIdx);
                        if (srcCell != null)
                        {
                            ICell destCell = destRow.CreateCell(cellIdx);
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
            ICell cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            // 赋值
            cell.SetCellValue(value);

        }
        //写日期
        private static void SetXlsxCellString(ISheet sheet, int rowIdx, int colIdx, string value)
        {
            // 获取或创建行
            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);
            // 获取或创建单元格
            ICell cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);
            // 赋值
            cell.SetCellValue(value);

        }

        public async Task<List<SourceData>> GetSourceData(DateTime StartTime, DateTime EndtTime)
        {
            var result = await _sourceData.GetByDataTimeAsync(StartTime, EndtTime );
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
            var targetData = await _sourceData.db
                .FirstOrDefaultAsync(d => d.createdtime >= targetDateTime
                                        && d.createdtime < targetDateTime.AddHours(1));

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
            _sourceData.Update(targetData); // 标记实体为修改状态
            await _dbContext.SaveChangesAsync(); // 提交到数据库

            return true;
        }

    }


}
