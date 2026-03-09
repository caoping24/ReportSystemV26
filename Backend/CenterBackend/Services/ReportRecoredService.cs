using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;
using Microsoft.EntityFrameworkCore;
using System.Reflection;
using static System.Runtime.InteropServices.JavaScript.JSType;

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


        //{
        //                // 当天
        //        DateTime startTime = queryDate.Date.AddHours(0);
        //DateTime endTime = startTime.AddHours(23).AddMinutes(59);

        //var calculatedDatas = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);

        //// 3.构建【带日期维度的分组键】
        //var dataWithKey = calculatedDatas.Select(cd => new
        //{
        //    Data = cd,
        //    GroupKey = cd.ReportedTime >= startTime && cd.ReportedTime < endTime
        //        ? cd.ReportedTime.Hour
        //        : cd.ReportedTime.Hour + 100
        //}).ToList();

        //// 4. 按唯一分组键分组（保留原有逻辑）
        //var hourGroupDict = dataWithKey
        //    .GroupBy(item => item.GroupKey)
        //    .ToDictionary(g => g.Key, g => g.FirstOrDefault()?.Data);

        //var hourList = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };

        //var hourDataList = hourList.Select((hour, index) =>// 6. 构建返回数据（核心修改：计算真实时间+动态判定IsNextDay）
        //{

        //    DateTime realHourTime;
        //    realHourTime = queryDate.Date.AddHours(hour);


        //    int targetKey = hour;
        //    hourGroupDict.TryGetValue(targetKey, out var targetData);


        //    // 未来时间：该时段的真实时间 大于 当前系统时间
        //    bool isFutureTime = realHourTime > DateTime.Now;

        //    // IsNextDay=true（前端禁用）：未来时间 OR 无对应数据
        //    // IsNextDay=false（前端可编辑）：过去时间 AND 有对应数据
        //    bool isNextDay = isFutureTime || targetData == null;

        //    // 初始化返回DTO，赋值核心字段
        //    var hourData = new HourDataDto
        //    {
        //        Hour = hour,
        //        Date = date,
        //        IsNextDay = isNextDay, // 赋值修正后的禁用标识
        //        Cells = new Dictionary<string, string>() // 确保Cells初始化，避免空引用
        //    };

        //    // 7. 填充Cell字段（保留原有格式化逻辑，无数据为空字符串）
        //    hourData.Cells["Cell29"] = targetData?.Cell29?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell30"] = targetData?.Cell30?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell31"] = targetData?.Cell31?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell32"] = targetData?.Cell32?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell33"] = targetData?.Cell33?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell34"] = targetData?.Cell34?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell35"] = targetData?.Cell35?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell56"] = targetData?.Cell56?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell57"] = targetData?.Cell57?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell58"] = targetData?.Cell58?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell59"] = targetData?.Cell59?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell60"] = targetData?.Cell60?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell82"] = targetData?.Cell82?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell83"] = targetData?.Cell83?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell84"] = targetData?.Cell84?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell85"] = targetData?.Cell85?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell86"] = targetData?.Cell86?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell87"] = targetData?.Cell87?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell135"] = targetData?.Cell135?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell136"] = targetData?.Cell136?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell137"] = targetData?.Cell137?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell138"] = targetData?.Cell138?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell139"] = targetData?.Cell139?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell140"] = targetData?.Cell140?.ToString("0.00") ?? "";
        //    hourData.Cells["Cell141"] = targetData?.Cell141?.ToString("0.00") ?? "";

        //    return hourData;
        //}).ToList();

        //}
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
