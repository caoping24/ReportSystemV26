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
        private readonly IReportUnitOfWork _reportUnitOfWork;
        //private readonly CenterReportDbContext _dbContext
        private readonly IDataViewToExcel _dataViewToExcel;
        private readonly IDataToViewService _dataToViewService;
        private readonly IFileServices _fileService;
        private readonly ICalculatedAndSaveService _calculatedAndSaveService;

        public ReportService(IReportRepository<SourceData> sourceData,
            IReportRepository<OperatorInputData> operatorInputData,
            IReportRecordRepository<ReportRecord> reportRecord,
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
                    DayWorkBook dayCollections = new()
                    {
                        SheetType = SheetType.DayReport,
                        ReportedTime = fileInfo.ReportedTime,
                        Directory = fileInfo.Directory,
                        FileName = fileInfo.FileName,
                        ModFilePath = fileInfo.ModFilePath,
                    };
                    if (await _dataToViewService.DayGetMapDataAsync(dayCollections))
                    {
                        isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(dayCollections);
                    }
                    break;
                case 2:
                    break;
                case 3: 
                    break;
                case 4:
                    WeekWorkBook weekDataCollections = new()
                    {
                        SheetType = SheetType.WeekReport,
                        ReportedTime = fileInfo.ReportedTime,
                        Directory = fileInfo.Directory,
                        FileName = fileInfo.FileName,
                        ModFilePath = fileInfo.ModFilePath,
                    };
                    if (await _dataToViewService.WeekGetMapDataAsync(weekDataCollections))
                    {
                        isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(weekDataCollections);
                    }
                    
                    break;
                default:
                    break;
            }
            if (!isBuildSuccess)
                return false;

            //更新或插入记录
            {
                var existingRecord = await _reportRecord.Db.AsQueryable()
                    .Where(r => r.ReportedTime.Date == fileInfo.ReportedTime.Date && r.Type == fileInfo.Type)
                    .FirstOrDefaultAsync();

                if (existingRecord != null)
                {
                    existingRecord.ReportedTime = fileInfo.ReportedTime.Date;
                    existingRecord.LastChange = DateTime.Now;
                }
                else
                {
                    existingRecord = new ReportRecord()//插入记录
                    {
                        ReportedTime = fileInfo.ReportedTime,
                        LastChange = DateTime.Now,
                        Type = fileInfo.Type
                    };
                    await _reportRecord.AddAsync(existingRecord);
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

            var reportInfo = new ReportInfo
            {
                TimeStart = dto.Time.Date.AddHours(8),//当日8点
                TimeEnd = dto.Time.Date.AddHours(8).AddDays(1),
                SheetType = SheetType.DayReport,
            };
            return await _calculatedAndSaveService.DataAnalyses(reportInfo);
        }


    }


}
