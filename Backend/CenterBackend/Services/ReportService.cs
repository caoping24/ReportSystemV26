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

        /// <summary>
        /// 根据传入的Type类型，计算对应维度的统计数据并插入到CalculatedData表中 注意传入的时间
        /// </summary>
        /// <param name="_Dto"></param>
        /// <returns></returns>
        public async Task<bool> DataAnalyses(CalculateAndInsertDto _Dto)
        {
            DateTime StartTime;
            DateTime StopTime;

            switch (_Dto.type)
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
            return await CalculatedDataAndInsert(StartTime, StopTime, _Dto.type);
        }

        /// <summary>
        /// 根据Tpye类型，计算周/月/年统计数据
        /// </summary>
        private async Task<bool> CalculatedDataAndInsert(DateTime startTime, DateTime stopTime, int type)
        {
            var ReportedTime = startTime.Date;//记录是那一天的数据
            var target = _calculatedDatas.Db.FirstOrDefault(r => r.Type == type && r.ReportedTime == ReportedTime);
            bool isNewRecord = (target == null);
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



        private bool RebuildReport()    
        {
            return true;
        }

    }


}
