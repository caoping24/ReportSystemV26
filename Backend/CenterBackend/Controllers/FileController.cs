using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

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
        public async Task<IActionResult> Test1(int type)
        {
            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo = filePathGenerator.GetByType(DateTime.Now, type);
            if (string.IsNullOrEmpty(fileInfo.FileName))//检查Type是否合法，是否能找到对应的文件路径和文件名
            {
                return new BadRequestObjectResult(new { success = false, msg = "无效的请求参数" });
            }

            try
            {
                _fileService.CreateFolder(fileInfo.Directory);

                return new OkObjectResult(new { success = true, msg = "操作成功" });
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"创建文件夹失败：{ex.Message}" });
            }

        }
        // 根绝日期获取对应的文件夹路径和文件名测试
        [HttpGet("GetFolderPathTest")]
        public async Task<IActionResult> Test3(int type)
        {
            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            PathAndName fileInfo = filePathGenerator.GetByType(DateTime.Now, type);
            if (string.IsNullOrEmpty(fileInfo.FileName))//检查Type是否合法，是否能找到对应的文件路径和文件名
            {
                return new BadRequestObjectResult(new { success = false, msg = "无效的请求参数" });
            }
            return Ok(fileInfo);
        }

        // 下载单个文件测试
        [HttpGet("DownloadSingleFile")]
        public async Task<IActionResult> DownloadSingleFile(String timeStr, int type)
        {
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

        // 下载单个ZIP文件测试
        [HttpGet("DownloadZipFile")]
        public async Task<IActionResult> DownloadZipFile(String timeStr, int type)
        {
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
