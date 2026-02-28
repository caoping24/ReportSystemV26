using CenterBackend.IServices;
using CenterBackend.Logging;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("[controller]")]
    [AllowAnonymous] // 允许匿名访问，跳过认证
    public class HelpController : ControllerBase
    {
        private readonly IFileServices _fileService;
        private readonly IWebHostEnvironment _webHostEnv;
        private readonly IAppLogger _logger;
        public HelpController(IFileServices fileService, IWebHostEnvironment webHostEnv, IAppLogger _IAppLogger)
        {
            this._fileService = fileService;
            this._webHostEnv = webHostEnv;
            this._logger = _IAppLogger;
        }


        [HttpGet("Version")]
        public async Task<IActionResult> Version()
        {
            return Ok(new
            {
                success = true,
                msg = "Ver:1.1.0.0 LastChange:2026-02-26",
                compileTimeLocal = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            });
        }
    }
}