using CenterBackend.IServices;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FilterConfigController : ControllerBase
    {
        private readonly IFilterConfigService _configService;

        public FilterConfigController(IFilterConfigService configService)
        {
            _configService = configService;
        }

        [HttpPost("refresh-filters")]
        public async Task<IActionResult> RefreshFilters()
        {
            var (success, count, error) = await _configService.ReloadAsync();

            if (success)
                return Ok(new { success = true, message = $"已加载 {count} 条筛选规则" });
            else
                return StatusCode(500, new { success = false, message = "刷新失败", error });
        }

        [HttpGet("filters")]
        public IActionResult GetFilters()
        {
            var snapshot = _configService.GetSnapshot();

            if (snapshot == null)
                return Ok(new { loaded = false, message = "尚未加载筛选配置" });

            return Ok(new
            {
                loaded = true,
                loadedAt = snapshot.LoadedAt,
                rules = snapshot.Configs.Select(kv => new
                {
                    id = kv.Value.Id,
                    kv.Value.FieldName,
                    kv.Value.MinValue,
                    kv.Value.MaxValue,
                    kv.Value.Comment,
                    kv.Value.IsActive
                })
            });
        }

        /// <summary>
        /// 更新单条筛选配置 → 写 DB → 刷新内存
        /// </summary>
        [HttpPost("update-config")]
        public async Task<IActionResult> UpdateConfig([FromBody] UpdateFilterConfigDto dto)
        {
            var (success, error) = await _configService.UpdateConfigAsync(
                dto.Id, dto.MinValue, dto.MaxValue, dto.Comment);

            if (success)
                return Ok(new { success = true, message = "已更新", loadedAt = _configService.GetSnapshot()?.LoadedAt });
            else
                return StatusCode(500, new { success = false, message = "更新失败", error });
        }
        [HttpGet("filter-enabled")]
        public IActionResult GetFilterEnabled()
        {
            return Ok(new { enabled = _configService.IsFilterEnabled });
        }

        [HttpPost("filter-enabled")]
        public async Task<IActionResult> SetFilterEnabled([FromBody] FilterEnabledDto dto)
        {
            await _configService.SetFilterEnabledAsync(dto.Enabled);
            return Ok(new { success = true, enabled = dto.Enabled });
        }

    }

    public class UpdateFilterConfigDto
    {
        public int Id { get; set; }
        public float? MinValue { get; set; }
        public float? MaxValue { get; set; }
        public string? Comment { get; set; }
    }
    public class FilterEnabledDto
    {
        public bool Enabled { get; set; }
    }
}