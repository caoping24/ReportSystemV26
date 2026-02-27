using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
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

        /// <summary>
        /// 下载大文件压缩包测试
        /// </summary>
        /// <returns></returns>2026-01-26
        /// <exception cref=""></exception>
        [HttpGet("ZipDownloadFile")]
        public async Task<IActionResult> DownloadZipFileBig(String timeStr, int type)
        {
            string zipFileName = "default.zip";
            var filePathGenerator = new FilePathGenerator(_webHostEnv);
            var fileInfo = filePathGenerator.GetDay(timeStr.ToDateTime());

            var sourceFolder = fileInfo.Directory;
            string tempFolder = filePathGenerator.TempDirectory;
                    try
                    {
                if (Directory.Exists(tempFolder)) { Directory.Delete(tempFolder, recursive: true); }// 直接删除整个目录及所有内容
                bool compressSuccess = _fileService.CompressFolderToZip(sourceFolder, tempFolder, zipFileName);//调用FileService 压缩文件夹为Zip包
                if (!compressSuccess)
                {
                    var msg = "压缩失败，文件不存在或被占用.";
                    await _logger.LogErrorAsync(msg);
                    return BadRequest(msg);
                }

                string tempZipFullPath = Path.Combine(tempFolder, zipFileName);//生成的压缩文件完整路径
                if (!System.IO.File.Exists(tempZipFullPath))
                {
                    var msg = "压缩成功，但未生成下载文件";
                    await _logger.LogErrorAsync(msg);
                    return BadRequest(msg);
                }
                var fileStream = new FileStream(tempZipFullPath, FileMode.Open, FileAccess.Read, FileShare.Read,
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
