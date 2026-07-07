using CenterBackend.IServices;
using CenterBackend.Models;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;

namespace CenterBackend.Services
{
    public class DataFilterService
    {
        private readonly IFilterConfigService _configService;
        private readonly IReportRepository<SourceData> _sourceDataRepo;

        public DataFilterService(
            IFilterConfigService configService,
            IReportRepository<SourceData> sourceDataRepo)
        {
            _configService = configService;
            _sourceDataRepo = sourceDataRepo;
        }

        /// <summary>
        /// 查询一段时间内的 SourceData，按配置过滤，不合格字段返回 null
        /// </summary>
        /// <param name="start">开始时间</param>
        /// <param name="end">结束时间</param>
        /// <param name="dataType">可选的 Type 过滤（预留）</param>
        public async Task<List<Dictionary<string, object?>>> GetFilteredDataAsync(
            DateTime start, DateTime end, int? dataType = null)
        {
            var snapshot = _configService.GetSnapshot();

            // Bug 2 修复：快照为 null → 返回空列表而非抛异常
            if (snapshot == null)
                return new List<Dictionary<string, object?>>();

            // 从仓储层查询原始数据
            List<SourceData> rawData;
            if (dataType.HasValue)
                rawData = await _sourceDataRepo.GetByDateTimeRangeAsync(start, end, dataType.Value);
            else
                rawData = await _sourceDataRepo.GetByDateTimeRangeAsync(start, end);

            var result = new List<Dictionary<string, object?>>(rawData.Count);

            foreach (var row in rawData)
            {
                var dict = new Dictionary<string, object?>(snapshot.Configs.Count + 2)
                {
                    ["Id"] = row.Id,
                    ["ReportedTime"] = row.ReportedTime
                };

                foreach (var (fieldName, config) in snapshot.Configs)
                {
                    if (!snapshot.Getters.TryGetValue(fieldName, out var getter))
                        continue;

                    var rawValue = getter(row);

                    // 合格 → 保留原始值；不合格 → null
                    dict[fieldName] = config.IsValid(rawValue) ? rawValue : null;
                }

                result.Add(dict);
            }

            return result;
        }
    }
}