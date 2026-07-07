using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.Filters;
using CenterReport.Repository;
using CenterReport.Repository.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.Formula.Functions;
using System.Data;
using System.Linq.Expressions;
using System.Reflection;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;

namespace CenterBackend.Services
{
    public class FilterConfigService(IServiceScopeFactory scopeFactory) : IFilterConfigService
    {
        private readonly IServiceScopeFactory _scopeFactory = scopeFactory;
        private volatile FilterSnapshot? _snapshot;//筛选规则快照

        // 预缓存 SourceData 所有 float? 属性，OrdinalIgnoreCase 容错大小写
        private static readonly Dictionary<string, PropertyInfo> _sourceDataProps =
            typeof(SourceData)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.PropertyType == typeof(float?))
                .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

        public FilterSnapshot? GetSnapshot() => _snapshot;
        private volatile bool _isEnabled = true;
        public bool IsFilterEnabled => _isEnabled;
        public async Task<(bool Success, int Count, string? Error)> ReloadAsync()
        {
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var ctx = scope.ServiceProvider.GetRequiredService<CenterReportDbContext>();

                var configs = await QueryConfigsAsync(ctx);
                var getters = BuildGetters(configs);
                _isEnabled = await QueryFilterEnabledAsync(ctx);
                var snapshot = new FilterSnapshot(configs, getters, DateTime.Now);

                // 原子替换，旧快照被 GC
                _snapshot = snapshot;

                return (true, configs.Count, null);
            }
            catch (Exception ex)
            {
                // 失败保留旧配置，仅返回错误信息
                return (false, 0, ex.Message);
            }
        }

        private static async Task<Dictionary<string, FieldRangeFilter>> QueryConfigsAsync(
            CenterReportDbContext ctx)
        {
            var result = new Dictionary<string, FieldRangeFilter>(StringComparer.OrdinalIgnoreCase);

            await using var cmd = ctx.Database.GetDbConnection().CreateCommand();
            cmd.CommandText = """
                SELECT Id, FieldName, Comment, MinValue, MaxValue,IsActive
                FROM dbo.FieldRangeFilters
                WHERE IsActive = 1
                """;

            await ctx.Database.OpenConnectionAsync();
            await using var reader = await cmd.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var fieldName = reader.GetString(1);
                result[fieldName] = new FieldRangeFilter
                {
                    Id = reader.GetInt32(0),
                    FieldName = fieldName,
                    Comment = reader.IsDBNull(2) ? null : reader.GetString(2),
                    MinValue = reader.IsDBNull(3) ? null : reader.GetFloat(3),
                    MaxValue = reader.IsDBNull(4) ? null : reader.GetFloat(4),
                    IsActive = reader.GetBoolean(5)
                };
            }

            // 交由 using scope 结束时关闭连接
            return result;
        }

        private static Dictionary<string, Func<SourceData, float?>> BuildGetters(
            Dictionary<string, FieldRangeFilter> configs)
        {
            var getters = new Dictionary<string, Func<SourceData, float?>>(
                StringComparer.OrdinalIgnoreCase);

            var param = Expression.Parameter(typeof(SourceData), "r");

            foreach (var fieldName in configs.Keys)
            {
                // Bug 3 修复：无效属性名跳过，不抛异常
                if (!_sourceDataProps.TryGetValue(fieldName, out var propInfo))
                    continue;

                var prop = Expression.Property(param, propInfo);
                var expr = Expression.Lambda<Func<SourceData, float?>>(prop, param);
                getters[fieldName] = expr.Compile();
            }

            return getters;
        }
        public async Task<(bool Success, string? Error)> UpdateConfigAsync(
            int id, float? minValue, float? maxValue, string? comment)
        {
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var ctx = scope.ServiceProvider.GetRequiredService<CenterReportDbContext>();

                await using var cmd = ctx.Database.GetDbConnection().CreateCommand();
                cmd.CommandText = """
            UPDATE dbo.FieldRangeFilters
            SET MinValue = @min, MaxValue = @max, Comment = @comment
            WHERE Id = @id
            """;

                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
                cmd.Parameters.Add(new SqlParameter("@min", SqlDbType.Real)
                { Value = minValue ?? (object)DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@max", SqlDbType.Real)
                { Value = maxValue ?? (object)DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@comment", SqlDbType.NVarChar, 128)
                { Value = comment ?? (object)DBNull.Value });

                await ctx.Database.OpenConnectionAsync();
                await cmd.ExecuteNonQueryAsync();

                await ReloadAsync();
                return (true, null);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
        public List<SourceData> GetFilteredData(List<SourceData> sourceData)
        {
            if (!_isEnabled)
                return sourceData;
            var snapshot = _snapshot;
            if (snapshot == null || sourceData == null || sourceData.Count == 0)
                return [];

            // 深拷贝：反射遍历所有属性
            var props = typeof(SourceData)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var result = new List<SourceData>(sourceData.Count);
            foreach (var row in sourceData)
            {
                var copy = new SourceData();
                foreach (var prop in props)
                    prop.SetValue(copy, prop.GetValue(row));
                result.Add(copy);
            }

            // 对副本做筛选（改副本不影响 EF 跟踪）
            foreach (var row in result)
            {
                foreach (var (fieldName, config) in snapshot.Configs)
                {
                    if (!snapshot.Getters.TryGetValue(fieldName, out var getter))
                        continue;
                    if (!config.IsValid(getter(row)))
                        typeof(SourceData).GetProperty(fieldName)?.SetValue(row, null);
                }
            }

            return result;
        }
        private static async Task<bool> QueryFilterEnabledAsync(CenterReportDbContext ctx)
        {
            await using var cmd = ctx.Database.GetDbConnection().CreateCommand();
            cmd.CommandText = "SELECT [Value] FROM dbo.FilterGlobalSettings WHERE [Key] = 'IsFilterEnabled'";

            await ctx.Database.OpenConnectionAsync();
            var result = await cmd.ExecuteScalarAsync();
            return result?.ToString() == "true";
        }
        public async Task SetFilterEnabledAsync(bool enabled)
{
    try
    {
        using var scope = _scopeFactory.CreateScope();
        var ctx = scope.ServiceProvider.GetRequiredService<CenterReportDbContext>();

        await using var cmd = ctx.Database.GetDbConnection().CreateCommand();
        cmd.CommandText = """
            UPDATE dbo.FilterGlobalSettings SET [Value] = @val WHERE [Key] = 'IsFilterEnabled'
            """;
        cmd.Parameters.Add(new SqlParameter("@val", SqlDbType.NVarChar, 16)
            { Value = enabled ? "true" : "false" });

        await ctx.Database.OpenConnectionAsync();
        await cmd.ExecuteNonQueryAsync();

        _isEnabled = enabled;
    }
    catch
    {
        // 写 DB 失败不抛，前端显示旧状态
    }
}
    }
}