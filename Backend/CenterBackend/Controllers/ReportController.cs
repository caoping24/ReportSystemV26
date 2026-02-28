using CenterBackend.common;
using CenterBackend.Dto;
using CenterBackend.Exceptions;
using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

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
            if (_CalculateAndInsertDto.type == 0)
            {
                throw new BusinessException(ErrorCode.PARAMS_ERROR, "类型错误");
            }
            var result = await reportService.DataAnalyses(_CalculateAndInsertDto);
            return ResultUtils<bool>.Success(result);
        }

        //  根据传入时间查询数据库,生成报表 Type 表示不同的报表类型
        [HttpPost("BuildReport")]
        public async Task<IActionResult> CreateAndBuildReport([FromBody] CreateReportDto createReportDto)
        {
            //await _logger.LogInfoAsync($"CreateAndBuildReport:CreateReportDto: {_CreateReportDto.type},{_CreateReportDto.Time}");


            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo = filePathGenerator.GetByType(createReportDto.Time, createReportDto.Type);
            if (string.IsNullOrEmpty(fileInfo.FileName))//检查Type是否合法，是否能找到对应的文件路径和文件名
            {
                return new BadRequestObjectResult(new { success = false, msg = "无效的请求参数" });
            }
            try
            {
                _fileService.CreateFolder(fileInfo.Directory);//自动创建文件夹
                return await reportService.WriteXlsxAndSave(fileInfo.ModFilePath, fileInfo.FullPath, createReportDto.Time, createReportDto.Type);
            }
            catch (Exception ex)
            {
                return BadRequest($"操作异常：{ex.Message}");
            }
        }
        /// <summary>
        /// 下载单个Excel文件
        /// </summary>
        /// <param name="loginDto"></param>
        /// <returns></returns>
        [HttpGet("DownloadExcel")]
        public async Task<IActionResult> DownloadFile(String timeStr, int type)
        {
            //await _logger.LogInfoAsync($"DownloadFile:timeStr:{timeStr},Type:{type}");
            bool isValid = DateTime.TryParseExact(
                timeStr,
                "yyyy-MM-dd HH:mm:ss",
                CultureInfo.InvariantCulture, // 替代null，避免区域设置影响（比如中文/英文系统）
                DateTimeStyles.None,
                out DateTime fileDate);
            if (!isValid)
            {
                return new BadRequestObjectResult(new { success = false, msg = "时间格式错误" });
            }
            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo = filePathGenerator.GetByType(fileDate, type);
            if (string.IsNullOrEmpty(fileInfo.FileName))//检查Type是否合法，是否能找到对应的文件路径和文件名
            {
                return new BadRequestObjectResult(new { success = false, msg = "无效的请求参数" });
            }
            try
            {
                var (filePath, contentType, downloadFileName) = _fileService.DownloadFileInfo(fileInfo.Directory, fileInfo.FileName);
                return PhysicalFile(filePath, contentType, downloadFileName);//官方推荐：直接用 PhysicalFile 自动处理文件流、响应头、范围请求（大文件下载）
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"{ex}" });
            }

        }
    }
}
