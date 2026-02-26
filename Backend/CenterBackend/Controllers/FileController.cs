using CenterBackend.IFileService;
using CenterBackend.Logging;
using CenterBackend.Models;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;
using System.Net.Mime;

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

        /// <summary>
        /// 在Report目录下创建按日期命名所有的文件夹
        /// </summary>
        /// <returns></returns>
        /// <exception cref=""></exception>
        [HttpGet("CreateDateFolderTest")]
        public async Task<IActionResult> Test1()
        {
            try
            {
                //var pathAndName = _fileService.GetDateFolderPathAndName(Path.Combine(_webHostEnv.ContentRootPath, "Report"), DateTime.Now);
                //string sourceFilePath = Path.Combine(_webHostEnv.WebRootPath, "Files/Model-20260116.xlsx");
                _fileService.CreateDateFolder(Path.Combine(_webHostEnv.ContentRootPath, "Report"), DateTime.Now);
                return new OkObjectResult(new { success = true, msg = "操作成功" });
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(new { success = false, msg = $"创建文件夹失败：{ex.Message}" });
            }
        }
        /// <summary>
        /// 根绝日期获取对应的文件夹路径和文件名测试
        /// </summary>
        /// <returns></returns>
        /// <exception cref=""></exception>
        [HttpGet("GetFolderPathTest")]
        public async Task<FilePathAndName> Test3()
        {
            await _logger.LogInfoAsync($"Test3:timeStr");
            var temp = _fileService.GetDateFolderPathAndName(Path.Combine(_webHostEnv.ContentRootPath, "Report"), DateTime.Now);
            return temp;

        }

        /// <summary>
        /// 下载大文件压缩包测试
        /// </summary>
        /// <returns></returns>2026-01-26
        /// <exception cref=""></exception>
        [HttpGet("ZipDownloadFile")]
        public async Task<IActionResult> DownloadZipFileBig(String timeStr, int type)
        {
            string zipFileName = "default.zip";
            string sourceFolder = Path.Combine(_webHostEnv.ContentRootPath, "Report");
            string tempZipPath = string.Empty;
            await _logger.LogInfoAsync($"DownloadZipFileBig:timeStr: {timeStr},type: {type}");
            try
            {
                // 日期解析容错：严格匹配yyyy-MM-dd
                if (!DateTime.TryParseExact(timeStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTime))
                {
                    var msg = $"日期格式无效.";
                    await _logger.LogErrorAsync(msg);
                    return BadRequest(msg);
                }
                switch (type)
                {
                    case 1://day
                        sourceFolder = Path.Combine(sourceFolder, "日报表", $"{dateTime.Year}-{dateTime.Month:00}");
                        zipFileName = $"日报表_{dateTime.Year}-{dateTime.Month:00}.zip";
                        break;
                    case 2://week
                        sourceFolder = Path.Combine(sourceFolder, "周报表", $"{dateTime.Year}-{dateTime.Month:00}");
                        zipFileName = $"周报表_{dateTime.Year}-{dateTime.Month:00}.zip";
                        break;
                    case 3://month
                        sourceFolder = Path.Combine(sourceFolder, "月报表", $"{dateTime.Year}");
                        zipFileName = $"月报表_{dateTime.Year}-{dateTime.Month:00}.zip";
                        break;
                    default:
                        return BadRequest("类型无效，请检查类型！");
                }
                string tempFolder = Path.Combine(_webHostEnv.WebRootPath, "Temp");

                if (Directory.Exists(tempFolder))//清理temp文件夹
                {
                    try
                    {
                        Directory.Delete(tempFolder, recursive: true);// 直接删除整个目录及所有内容
                    }
                    catch (Exception ex)
                    {

                        await _logger.LogErrorAsync($"删除Temp目录失败：{ex.Message}");
                    }
                }
                tempZipPath = Path.Combine(_webHostEnv.WebRootPath, "Temp", zipFileName);// 服务器临时压缩包路径
                string tempDir = Path.GetDirectoryName(tempZipPath);
                Directory.CreateDirectory(tempDir);
                bool compressSuccess = _fileService.CompressFolderToZip(sourceFolder, tempZipPath);//调用FileService 压缩文件夹为Zip包
                if (!compressSuccess)
                {
                    var msg = "压缩失败，文件不存在或被占用.";
                    await _logger.LogErrorAsync(msg);
                    return BadRequest(msg);
                }
                if (!System.IO.File.Exists(tempZipPath))
                {
                    var msg = "压缩成功，但未生成下载文件";
                    await _logger.LogErrorAsync(msg);
                    return BadRequest(msg);
                }
                var fileStream = new FileStream(tempZipPath, FileMode.Open, FileAccess.Read, FileShare.Read,
                    bufferSize: 4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
                await _logger.LogInfoAsync($"下载Zip文件成功.");
                return File(fileStream, MediaTypeNames.Application.Zip, zipFileName);// 流式返回，ASP.NET Core自动管理流释放
            }
            catch (Exception ex)
            {
                var msg = $"下载失败：{ex.Message}";
                await _logger.LogErrorAsync(msg);
                return BadRequest(msg);
            }
        }
    }
}
