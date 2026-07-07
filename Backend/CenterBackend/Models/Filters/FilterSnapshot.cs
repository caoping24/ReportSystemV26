using CenterReport.Repository.Models;

namespace CenterBackend.Models.Filters
{
    public class FilterSnapshot
    {
        public Dictionary<string, FieldRangeFilter> Configs { get; }
        public Dictionary<string, Func<SourceData, float?>> Getters { get; }
        public DateTime LoadedAt { get; }

        public FilterSnapshot(
            Dictionary<string, FieldRangeFilter> configs,
            Dictionary<string, Func<SourceData, float?>> getters,
            DateTime loadedAt)
        {
            Configs = configs ?? throw new ArgumentNullException(nameof(configs));
            Getters = getters ?? throw new ArgumentNullException(nameof(getters));
            LoadedAt = loadedAt;
        }
    }
}