using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.CalculateData;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.HPSF;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.Drawing;
using System.Reflection;
using System.Security.Cryptography;

namespace CenterBackend.Services
{
    public class ReportService : IReportService
    {
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData;
        private readonly IReportRecordRepository<ReportRecord> _reportRecord;
        private readonly IReportRepository<CalculatedData> _calculatedDatas;
        private readonly IReportUnitOfWork _reportUnitOfWork;
        //private readonly CenterReportDbContext _dbContext
        private readonly IDataViewToExcel _dataViewToExcel;
        private readonly IDataToViewService _dataToViewService;
        private readonly IFileServices _fileService;
        private readonly ICalculatedAndSaveService _calculatedAndSaveService;

        public ReportService(IReportRepository<SourceData> sourceData,
            IReportRepository<OperatorInputData> operatorInputData,
            IReportRecordRepository<ReportRecord> reportRecord,
            IReportRepository<CalculatedData> CalculatedDatas,
            IReportUnitOfWork reportUnitOfWork,
            //IHttpContextAccessor httpContextAccessor,
            //CenterReportDbContext _dbContext,
            IDataViewToExcel dataViewToExcel,
            IDataToViewService dataToViewService,
            IFileServices fileService,
            ICalculatedAndSaveService calculatedAndSaveService
            )
        {
            this._sourceData = sourceData;
            this._operatorInputData = operatorInputData;
            this._reportRecord = reportRecord;
            this._calculatedDatas = CalculatedDatas;
            this._reportUnitOfWork = reportUnitOfWork;
            this._dataViewToExcel = dataViewToExcel;
            //this._dbContext = _dbContext;
            this._dataToViewService = dataToViewService;
            this._fileService = fileService;
            this._calculatedAndSaveService = calculatedAndSaveService;

        }

        public async Task<bool> RebuildReport(PathAndName fileInfo)
        {
            bool isBuildSuccess = false;
            switch (fileInfo.Type)
            {
                case 1://日报表
                    DayWorkBook datacollections = new()
                    {
                        SheetType = SheetType.DayReport,
                        ReportedTime = fileInfo.ReportedTime,
                        Directory = fileInfo.Directory,
                        FileName = fileInfo.FileName,
                        ModFilePath = fileInfo.ModFilePath,
                    };
                    if (await _dataToViewService.DayGetMapDataAsync(datacollections))
                    {
                        isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(datacollections);
                    }
                        break;
                case 2:
                    break;
                default:
                    break;
            }
            if (!isBuildSuccess)
                return false;

            //更新或插入记录
            {
                var existingRecord = await _reportRecord.db.AsQueryable()
                    .Where(r => r.ReportedTime.Date == fileInfo.ReportedTime.Date && r.Type == fileInfo.Type)
                    .FirstOrDefaultAsync();

                if (existingRecord != null)
                {
                    existingRecord.ReportedTime = fileInfo.ReportedTime.Date;
                    existingRecord.LastChange = DateTime.Now;
                }
                else
                {
                    ReportRecord reportRecord = new()//插入记录
                    {
                        ReportedTime = fileInfo.ReportedTime,
                        LastChange = DateTime.Now,
                        Type = fileInfo.Type
                    };
                    await _reportRecord.AddAsync(reportRecord);
                }
                await _reportUnitOfWork.SaveChangesAsync();
            }
            return true;
        }

        /// <summary>
        /// 根据传入的Type类型，计算对应维度的统计数据并插入到CalculatedData表中 注意传入的时间
        /// </summary>
        /// <param name="_Dto"></param>
        /// <returns></returns>
        public async Task<bool> ConfigDataAnalyses(CalculateAndInsertDto dto)
        {

            var reportInfo = new ReportInfo(); 
            switch (dto.Type)
            {
                case 1: // 昨天
                    reportInfo.TimeStart = dto.Time.Date.AddHours(8);//当日8点
                    reportInfo.TimeEnd = reportInfo.TimeStart.AddDays(1);
                    reportInfo.SheetType = SheetType.DayReport;
                    break;
                case 2: // 上月
                    reportInfo.TimeStart = new DateTime(dto.Time.Year, dto.Time.Month, 1).AddMonths(-1);// 计算上月的开始时间（1号）
                    reportInfo.TimeEnd = reportInfo.TimeStart.AddDays(7);
                    reportInfo.SheetType = SheetType.MonthReport;
                    break;
                case 3: // 去年
                    reportInfo.TimeStart = new DateTime(dto.Time.Year, 1, 1).AddYears(-1);// 计算去年的开始时间（1月1号）
                    reportInfo.TimeEnd = new DateTime(dto.Time.Year, 1, 1).AddDays(-1); // 去年的结束时间（12月31号）
                    reportInfo.SheetType = SheetType.YearReport;
                    break;
                case 4: // 上周
                    DateTime currentDayOfWeek = dto.Time.Date;// 计算上周的开始时间（星期一）
                    int daysToLastMonday = ((int)currentDayOfWeek.DayOfWeek + 6) % 7 + 7;

                    reportInfo.TimeStart = currentDayOfWeek.AddDays(-daysToLastMonday);
                    reportInfo.TimeEnd = reportInfo.TimeStart.AddDays(7);
                    reportInfo.SheetType = SheetType.WeekReport;
                    break;
                default:
                    return false;
            }

            return await _calculatedAndSaveService.DataAnalyses(reportInfo);
        }


    }


}
