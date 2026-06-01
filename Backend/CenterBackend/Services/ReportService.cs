using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using Microsoft.EntityFrameworkCore;

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

        public ReportService(IReportRepository<SourceData> sourceData,
            IReportRepository<OperatorInputData> operatorInputData,
            IReportRecordRepository<ReportRecord> reportRecord,
            IReportUnitOfWork reportUnitOfWork,
            //IHttpContextAccessor httpContextAccessor,
            //CenterReportDbContext _dbContext,
            IDataViewToExcel dataViewToExcel,
            IDataToViewService dataToViewService,
            IFileServices fileService
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
                        var FilePath = System.IO.Path.Combine(dayCollections.Directory, dayCollections.FileName);
                        if (_fileService.CopyFile(dayCollections.ModFilePath, FilePath))
                            isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(dayCollections);
                    }
                    break;
                case 2:
                    MonthWorkBook monthDataCollections = new()
                    {
                        SheetType = SheetType.MonthReport,
                        ReportedTime = fileInfo.ReportedTime,
                        Directory = fileInfo.Directory,
                        FileName = fileInfo.FileName,
                        ModFilePath = fileInfo.ModFilePath,
                    };
                    if (await _dataToViewService.MonthGetMapDataAsync(monthDataCollections))
                    {
                        var FilePath = System.IO.Path.Combine(monthDataCollections.Directory, monthDataCollections.FileName);
                        if (_fileService.CopyFile(monthDataCollections.ModFilePath, FilePath))
                            isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(monthDataCollections);
                    }
                    break;
                case 3:
                    YearWorkBook yearDataCollections = new()
                    {
                        SheetType = SheetType.YearReport,
                        ReportedTime = fileInfo.ReportedTime,
                        Directory = fileInfo.Directory,
                        FileName = fileInfo.FileName,
                        ModFilePath = fileInfo.ModFilePath,
                    };
                    if (await _dataToViewService.YearGetMapDataAsync(yearDataCollections))
                    {
                        var FilePath = System.IO.Path.Combine(yearDataCollections.Directory, yearDataCollections.FileName);
                        if (_fileService.CopyFile(yearDataCollections.ModFilePath, FilePath))
                            isBuildSuccess = await _dataViewToExcel.WriteXlsxAndSaveAsync(yearDataCollections);
                    }
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
                        var FilePath = System.IO.Path.Combine(weekDataCollections.Directory, weekDataCollections.FileName);
                        if (_fileService.CopyFile(weekDataCollections.ModFilePath, FilePath))
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




    }


}
