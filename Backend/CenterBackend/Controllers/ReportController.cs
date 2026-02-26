using CenterBackend.common;
using CenterBackend.Dto;
using CenterBackend.Exceptions;
using CenterBackend.IFileService;
using CenterBackend.IReportServices;
using CenterBackend.Logging;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly IReportService reportService;
        private readonly IFileServices _fileService;
        private readonly IWebHostEnvironment _webHostEnv;
        private readonly IAppLogger _logger;

        public ReportController(IReportService reportService, IFileServices fileService, IWebHostEnvironment webHostEnv, IAppLogger _IAppLogger)
        {
            this.reportService = reportService;
            this._fileService = fileService;
            this._webHostEnv = webHostEnv;
            this._logger = _IAppLogger;
        }
        /// <summary>
        /// 根据dto.Type 统计数据并且插入表中
        /// </summary>
        /// <param name="_AddReportDailyDto"></param>
        /// <returns></returns>
        /// <exception cref="BusinessException"></exception>
        [HttpPost("AnalysesInsert")]
        public async Task<BaseResponse<bool>> AnalysesInsert([FromBody] CalculateAndInsertDto _CalculateAndInsertDto)
        {
            await _logger.LogInfoAsync($"AnalysesInsert:CalculateAndInsertDto: {_CalculateAndInsertDto.Time},{_CalculateAndInsertDto.Time}");
            if (_CalculateAndInsertDto.Type == 0)
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "类型错误");
            }
            var result = await reportService.DataAnalyses(_CalculateAndInsertDto);
            return ResultUtils<bool>.Success(result);
        }
        /// <summary>
        ///  根据传入时间查询数据库,生成报表 Type 表示不同的报表类型
        /// </summary>
        /// <param name="CreateReportDto"></param>
        /// <returns></returns>
        [HttpPost("BuildReport")]
        public async Task<IActionResult> CreateAndBuildReport([FromBody] CreateReportDto _CreateReportDto)
        {
            await _logger.LogInfoAsync($"CreateAndBuildReport:CreateReportDto: {_CreateReportDto.Type},{_CreateReportDto.Time}");

            var reportFileRoot = Path.Combine(_webHostEnv.ContentRootPath, "Report");//所有报表汇总文件夹
            DateTime tempTime = _CreateReportDto.Time;
            int reportType = _CreateReportDto.Type;

            var filePathAndName = _fileService.GetDateFolderPathAndName(reportFileRoot, tempTime);
            if (string.IsNullOrWhiteSpace(filePathAndName.DailyFileName)) return BadRequest("获取文件路径失败，请检查传入日期格式！");
            string? XlsxFilesPath;
            string? XlsxFilesFullPath;
            string? modelFilePath;
            try
            {
                switch (_CreateReportDto.Type)
                {
                    case 1: //Daily
                        XlsxFilesPath = filePathAndName.DailyFilesPath;
                        XlsxFilesFullPath = filePathAndName.DailyFilesFullPath;
                        modelFilePath = Path.Combine(_webHostEnv.WebRootPath, "Model\\Model-20260116.xlsx");//日报表模板路径
                        break;
                    case 2: //Weekly
                        XlsxFilesPath = filePathAndName.WeeklyFilesPath;
                        XlsxFilesFullPath = filePathAndName.WeeklyFilesFullPath;
                        modelFilePath = Path.Combine(_webHostEnv.WebRootPath, "Model\\Model-20251208-Week.xlsx");
                        break;
                    case 3: //Monthly
                        XlsxFilesPath = filePathAndName.MonthlyFilesPath;
                        XlsxFilesFullPath = filePathAndName.MonthlyFilesFullPath;
                        modelFilePath = Path.Combine(_webHostEnv.WebRootPath, "Model\\Model-20260116.xlsx");
                        break;
                    case 4: //Yearly
                        XlsxFilesPath = filePathAndName.YearlyFilesPath;
                        XlsxFilesFullPath = filePathAndName.YearlyFilesFullPath;
                        modelFilePath = Path.Combine(_webHostEnv.WebRootPath, "Model\\Model-20260116.xlsx");
                        break;
                    default:
                        return new BadRequestObjectResult(new { success = false, msg = "ReportType不存在" });
                }
                _fileService.CreateFolder(XlsxFilesPath);//自动创建文件夹
                return await reportService.WriteXlsxAndSave(modelFilePath, XlsxFilesFullPath, tempTime, reportType);
            }
            catch (Exception ex)
            {
                return BadRequest($"操作异常：{ex.Message}");
            }
        }
        /// <summary>
        /// 下载单个excle文件
        /// </summary>
        /// <param name="loginDto"></param>
        /// <returns></returns>
        [HttpGet("DownloadExcel")]
        public async Task<IActionResult> DownloadFile(String timeStr, int Type)
        {
            await _logger.LogInfoAsync($"DownloadFile:timeStr:{timeStr},Type:{Type}");

            var modelFilePath = Path.Combine(_webHostEnv.ContentRootPath, "Report");//日报表模板路径

            DateTime dateTime = DateTime.ParseExact(timeStr, "yyyy-MM-dd HH:mm:ss", null);
            var PathAndFileName = _fileService.GetDateFolderPathAndName(modelFilePath, dateTime);
            var DownloadfilePath = string.Empty;
            var DownloadfileName = string.Empty;
            string fileName;
            switch (Type)
            {
                case 1: //Daily
                    DownloadfilePath = PathAndFileName.DailyFilesPath;
                    DownloadfileName = PathAndFileName.DailyFileName;
                    fileName = PathAndFileName.DailyFileName;
                    break;
                case 2: //Weekly
                    DownloadfilePath = PathAndFileName.WeeklyFilesPath;
                    DownloadfileName = PathAndFileName.WeeklyFileName;
                    fileName = PathAndFileName.WeeklyFileName;
                    break;
                case 3: //Monthly
                    DownloadfilePath = PathAndFileName.MonthlyFilesPath;
                    DownloadfileName = PathAndFileName.MonthlyFileName;
                    fileName = PathAndFileName.MonthlyFileName;
                    break;
                case 4: //Yearly
                    DownloadfilePath = PathAndFileName.YearlyFilesPath;
                    DownloadfileName = PathAndFileName.YearlyFileName;
                    fileName = PathAndFileName.YearlyFileName;
                    break;
                default:
                    return BadRequest("类型错误，请检查传入类型！");
            }
            var (fileStream, encodeFileName) = _fileService.DownloadSingleFile(DownloadfilePath, DownloadfileName);
            if (fileStream == null)
            {
                return NotFound("文件不存在。");
            }

            Response.Headers.Append("Content-Disposition", $"attachment;filename={Uri.EscapeDataString(fileName)}");
            return File(fileStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
