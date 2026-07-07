using CenterBackend.Models;
using CenterBackend.Models.Filters;
using CenterReport.Repository.Models;

namespace CenterBackend.IServices
{
    public interface IFilterConfigService
    {
        /// <summary>获取当前快照（可能为 null）</summary>
        FilterSnapshot? GetSnapshot();

        /// <summary>重新加载配置 → 编译 Getter → 原子替换快照</summary>
        Task<(bool Success, int Count, string? Error)> ReloadAsync();
        /// <summary>更新单条配置 → 写 DB → 刷新内存</summary>
        Task<(bool Success, string? Error)> UpdateConfigAsync(
            int id, float? minValue, float? maxValue, string? comment);
        List<SourceData> GetFilteredData(List<SourceData> sourceData);
        bool IsFilterEnabled { get; }
        Task SetFilterEnabledAsync(bool enabled);
    }
}