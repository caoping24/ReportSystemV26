using CenterBackend.common;
using CenterBackend.Dto;
using CenterBackend.Exceptions;
using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
using CenterBackend.Models.CalculateData;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using SharpCompress.Common;
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


        //  根据传入时间查询数据库,生成报表 Type 表示不同的报表类型
        [HttpPost("BuildReport")]
        public async Task<IActionResult> BuildAndDownloadReport([FromBody] CreateReportDto createReportDto)
        {
            //await _logger.LogInfoAsync($"CreateAndBuildReport:CreateReportDto: {_CreateReportDto.type},{_CreateReportDto.Time}");

            var fileDate = createReportDto.Time;
            var type = createReportDto.Type;

            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo = filePathGenerator.GetByType(fileDate, type);
            if (string.IsNullOrEmpty(fileInfo.FileName))//检查Type是否合法，是否能找到对应的文件路径和文件名
            {
                return new BadRequestObjectResult(new { success = false, msg = "无效的请求参数" });
            }
            try
            {
                fileInfo.ReportedTime = createReportDto.Time;
                var isSuccess = await reportService.RebuildReport(fileInfo);
                if (!isSuccess)
                    return new BadRequestObjectResult(new { success = false, msg = "生成文件失败" });
                var (filePath, contentType, downloadFileName) = _fileService.DownloadFileInfo(fileInfo.Directory, fileInfo.FileName);
                return PhysicalFile(filePath, contentType, downloadFileName);//官方推荐：直接用 PhysicalFile 自动处理文件流、响应头、范围请求（大文件下载）
            }
            catch (Exception ex)
            {
                return BadRequest($"操作异常：{ex.Message}");
            }
        }



    }
}
