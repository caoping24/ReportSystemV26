using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
using CenterBackend.Services;
using Masuit.Tools;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using System.Globalization;
using System.Net.Mime;
using static FastExpressionCompiler.ExpressionCompiler;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FileController : ControllerBase
    {
        private readonly IFileServices _fileService;
        private readonly IWebHostEnvironment _webHostEnv;
        private readonly IAppLogger _logger;
        public FileController(IFileServices fileService, IWebHostEnvironment webHostEnv, IAppLogger _IAppLogger)
        {
            this._fileService = fileService;
            this._webHostEnv = webHostEnv;
            this._logger = _IAppLogger;
        }

        // 创建文件夹测试
        [HttpGet("CreateFolderTest")]
        public async Task<IActionResult> Test1()
        {
            var temp = new FilePathGenerator(_webHostEnv);
            PathAndName result;

            try
            {
                result = temp.GetDay(DateTime.Now);
                _fileService.CreateFolder(result.Directory);

                result = temp.GetMonth(DateTime.Now);
                _fileService.CreateFolder(result.Directory);

                result = temp.GetYear(DateTime.Now);
                _fileService.CreateFolder(result.Directory);

                result = temp.GetWeek(DateTime.Now);
                _fileService.CreateFolder(result.Directory);

                return new OkObjectResult(new { success = true, msg = "操作成功" });
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"创建文件夹失败：{ex.Message}" });
            }

        }
        // 根绝日期获取对应的文件夹路径和文件名测试
        [HttpGet("GetFolderPathTest")]
        public async Task<IActionResult> Test3(int id)
        {
            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo;
            switch (id)
            {
                case 1:
                    fileInfo = filePathGenerator.GetDay(DateTime.Now);
                    break;
                case 2:
                    fileInfo = filePathGenerator.GetMonth(DateTime.Now);
                    break;
                case 3:
                    fileInfo = filePathGenerator.GetYear(DateTime.Now);
                    break;
                case 4:
                    fileInfo = filePathGenerator.GetWeek(DateTime.Now);
                    break;
                default:
                    return new BadRequestObjectResult(new { error = "无效的请求参数" });
            }
            return Ok(fileInfo);
        }

        // 下载单个文件测试
        [HttpGet("DownloadSingleFile")]
        public async Task<IActionResult> DownloadSingleFile(String timeStr )
        {
            var filePathGenerator = new FilePathGenerator(_webHostEnv);

            PathAndName fileInfo = filePathGenerator.GetDay(timeStr.ToDateTime());

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

        // 下载单个ZIP文件测试
        [HttpGet("DownloadZipFile")]
        public async Task<IActionResult> DownloadZipFile(String timeStr)
        {
            var filePathGenerator = new FilePathGenerator(_webHostEnv);

            PathAndName fileInfo = filePathGenerator.GetDay(timeStr.ToDateTime());

            try
            {
                bool isSuccess = _fileService.CompressFolderToZip(fileInfo.Directory, filePathGenerator.TempDirectory, "测试.zip");
                if (isSuccess)
                {
                    var (filePath, contentType, downloadFileName) = _fileService.DownloadFileInfo(filePathGenerator.TempDirectory, "测试.zip");
                    return PhysicalFile(filePath, contentType, downloadFileName);//官方推荐：直接用 PhysicalFile 自动处理文件流、响应头、范围请求（大文件下载）
                }
                return new BadRequestObjectResult(new { success = false, msg = "下载失败" });
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"{ex}" });
            }
        }
    }
}
