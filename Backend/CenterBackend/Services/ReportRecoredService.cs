using CenterBackend.IServices;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;
using Microsoft.EntityFrameworkCore;
using System.Reflection;

namespace CenterBackend.Services
{
    public class ReportRecordService : IReportRecordService
    {
        private readonly IReportRecordRepository<ReportRecord> _reportRecord;
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly CenterReportDbContext _dbContext;

        public ReportRecordService(IReportRecordRepository<ReportRecord> reportRecord, IReportRepository<SourceData> sourceData, CenterReportDbContext _dbContext)
        {
            this._reportRecord = reportRecord;
            this._sourceData = sourceData;
            this._dbContext = _dbContext;
        }


        public async Task<PaginationResult<ReportRecord>> GetReportsByPageAsync(PaginationRequest request)
        {
            return await _reportRecord.GetReportByPageAsync(request);
        }
        public async Task<bool> UpdateSourceDataFieldAsync(string dateStr, int hour, string prop, string valueStr)
        {

            if (!DateTime.TryParse(dateStr, out DateTime targetDate))
            {
                throw new ArgumentException($"日期格式错误，要求yyyy-MM-dd，当前值：{dateStr}", nameof(dateStr));
            }

            // 构建目标时间（精确到小时，用于筛选记录）
            DateTime targetDateTime = new DateTime(targetDate.Year, targetDate.Month, targetDate.Day, hour, 0, 0);

            // 转换值为float?类型
            if (!float.TryParse(valueStr, out float value))
            {
                throw new ArgumentException($"值转换失败，要求浮点数字符串，当前值：{valueStr}", nameof(valueStr));
            }
            float? targetValue = value; // 兼容nullable float类型

            // 校验字段名是否存在
            PropertyInfo? propInfo = typeof(SourceData).GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
            if (propInfo == null)
            {
                throw new ArgumentException($"SourceData 不存在字段：{prop}", nameof(prop));
            }
            if (propInfo.PropertyType != typeof(float?))
            {
                throw new ArgumentException($"字段{prop}类型不是float?，不支持修改", nameof(prop));
            }
            // 方式1：精确匹配时间（可根据业务调整为时间范围）
            var targetData = await _sourceData.Db
                .FirstOrDefaultAsync(d => d.ReportedTime >= targetDateTime
                                        && d.ReportedTime < targetDateTime.AddHours(1));

            if (targetData == null)
            {
                throw new KeyNotFoundException($"未找到{targetDateTime:yyyy-MM-dd HH:mm}时间段的 SourceData 记录");
            }
            try
            {
                propInfo.SetValue(targetData, targetValue);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"设置字段{prop}值失败：{ex.Message}", ex);
            }
            await _sourceData.Update(targetData); // 标记实体为修改状态
            await _dbContext.SaveChangesAsync(); // 提交到数据库

            return true;
        }

    }
}
