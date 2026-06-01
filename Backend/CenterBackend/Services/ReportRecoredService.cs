using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using System.Reflection;

namespace CenterBackend.Services
{
    public class ReportRecordService : IReportRecordService
    {
        private readonly IReportRecordRepository<ReportRecord> _reportRecord;
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IOperatorInputDataRepository<OperatorInputData> _operatorInputData;
        private readonly CenterReportDbContext _dbContext;

        public ReportRecordService(IReportRecordRepository<ReportRecord> reportRecord, IReportRepository<SourceData> sourceData,
                                                    CenterReportDbContext _dbContext, IOperatorInputDataRepository<OperatorInputData> operatorInputData)

        {
            this._reportRecord = reportRecord;
            this._sourceData = sourceData;
            this._dbContext = _dbContext;
            this._operatorInputData = operatorInputData;
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

        //// 4. 按唯一分组键分组(保留原有逻辑)
        //var hourGroupDict = dataWithKey
        //    .GroupBy(item => item.GroupKey)
        //    .ToDictionary(g => g.Key, g => g.FirstOrDefault()?.Data);

        //var hourList = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };

        //var hourDataList = hourList.Select((hour, index) =>// 6. 构建返回数据(核心修改：计算真实时间+动态判定IsNextDay)
        //{

        //    DateTime realHourTime;
        //    realHourTime = queryDate.Date.AddHours(hour);


        //    int targetKey = hour;
        //    hourGroupDict.TryGetValue(targetKey, out var targetData);


        //    // 未来时间：该时段的真实时间 大于 当前系统时间
        //    bool isFutureTime = realHourTime > DateTime.Now;

        //    // IsNextDay=true(前端禁用)：未来时间 OR 无对应数据
        //    // IsNextDay=false(前端可编辑)：过去时间 AND 有对应数据
        //    bool isNextDay = isFutureTime || targetData == null;

        //    // 初始化返回DTO，赋值核心字段
        //    var hourData = new HourDataDto
        //    {
        //        Hour = hour,
        //        Date = date,
        //        IsNextDay = isNextDay, // 赋值修正后的禁用标识
        //        Cells = new Dictionary<string, string>() // 确保Cells初始化，避免空引用
        //    };

        //    // 7. 填充Cell字段(保留原有格式化逻辑，无数据为空字符串)
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

            // 构建目标时间(精确到小时，用于筛选记录)
            DateTime targetDateTime = new DateTime(targetDate.Year, targetDate.Month, targetDate.Day, hour, 0, 0);

            // 转换值为float?类型
            float? targetValue;
            if (string.IsNullOrEmpty(valueStr))
            {
                targetValue = null;
            }
            else if (!float.TryParse(valueStr, out float value))
            {
                throw new ArgumentException($"值转换失败，要求浮点数字符串或空字符串，当前值：{valueStr}", nameof(valueStr));
            }
            else
            {
                targetValue = value;
            }
            // 校验字段名是否存在
            PropertyInfo? propInfo = typeof(OperatorInputData).GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
            if (propInfo == null)
            {
                throw new ArgumentException($"OperatorInputData 不存在字段：{prop}", nameof(prop));
            }
            if (propInfo.PropertyType != typeof(float?))
            {
                throw new ArgumentException($"字段{prop}类型不是float?，不支持修改", nameof(prop));
            }
            // 方式1：精确匹配时间(可根据业务调整为时间范围)
            var targetData = await _operatorInputData.Db
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

            await _operatorInputData.Update(targetData); // 标记实体为修改状态
            await _dbContext.SaveChangesAsync(); // 提交到数据库

            return true;
        }

        public async Task<List<HourDataDto>> getHourDataTableOne(string date, string type)
        {
            // 1. 安全解析日期(仅保留日期部分，排除时分秒干扰)
            if (!DateTime.TryParse(date, out var targetDate))
            {
                throw new ArgumentException("日期格式无效，请传入如 '2026-03-10' 格式的日期", nameof(date));
            }
            targetDate = targetDate.Date; // 确保只取年月日

            // 2. 查询数据库中当天已存在的数据
            var existingData = await _operatorInputData.Db
                .Where(d => d.ReportedTime.Date == targetDate)
                .ToListAsync();

            // 3. 找出0-23点中缺失的小时数
            var existingHours = existingData.Select(d => d.ReportedTime.Hour).ToHashSet();
            var missingHours = Enumerable.Range(0, 24)
                                         .Where(hour => !existingHours.Contains(hour))
                                         .ToList();

            // 4. 为缺失的小时创建默认数据并批量插入数据库
            if (missingHours.Any())
            {
                var defaultEntities = missingHours.Select(hour => new OperatorInputData
                {
                    ReportedTime = targetDate.AddHours(hour),
                    // CreateTime = DateTime.Now
                }).ToList();

                _dbContext.OperatorInputDatas.AddRange(defaultEntities);
                await _dbContext.SaveChangesAsync(); // 提交数据库

                // 插入完成后，重新查询当天完整数据(核心修改点)
                // 此时数据库已包含原有数据 + 新增默认数据
                existingData = await _operatorInputData.Db
                    .Where(d => d.ReportedTime.Date == targetDate)
                    .ToListAsync();

            }
            return GetTableTypeOne(existingData, date);
        }

        public List<HourDataDto> GetTableTypeOne(List<OperatorInputData> operatorInputDatas, string date)
        {
            // 1. 校验日期格式(和原方法保持一致)
            if (!DateTime.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var queryDate))
            {
                throw new ArgumentException("日期格式错误，请传入YYYY-MM-DD格式", nameof(date));
            }

            // 2. 按上报时间的小时分组(核心：匹配原方法的小时分组逻辑)
            var hourGroupDict = operatorInputDatas
                .Where(data => data.ReportedTime.Date == queryDate.Date) // 仅保留查询日期的数据
                .GroupBy(data => data.ReportedTime.Hour) // 按小时分组
                .ToDictionary(g => g.Key, g => g.FirstOrDefault()); // 每个小时取第一条数据

            // 3. 定义小时列表(0-23)
            var hourList = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };

            // 4. 构建小时数据列表(核心逻辑和原方法对齐)
            var hourDataList = hourList.Select(hour =>
            {
                // 计算该小时的真实时间
                DateTime realHourTime = queryDate.Date.AddHours(hour);

                // 获取当前小时的对应数据
                hourGroupDict.TryGetValue(hour, out var targetData);

                // 判断是否为未来时间(原方法逻辑)
                bool isFutureTime = realHourTime > DateTime.Now;

                // 判定是否禁用(IsNextDay)：未来时间 或 无对应数据
                bool isNextDay = isFutureTime || targetData == null;

                // 初始化DTO并赋值核心字段
                var hourData = new HourDataDto
                {
                    Hour = hour,
                    Date = date,
                    IsNextDay = isNextDay,
                    Cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) // 忽略大小写，兼容前端
                };

                // 5. 【修复版】通过反射动态填充所有Cell1-Cell150字段
                for (int cellNum = 1; cellNum <= 150; cellNum++)
                {
                    // 获取模型中对应的Cell属性
                    PropertyInfo? cellProp = typeof(OperatorInputData).GetProperty($"Cell{cellNum}");
                    string cellKey = $"Cell{cellNum}"; // 前端使用的key
                    string cellValue = ""; // 默认空字符串

                    if (cellProp != null && targetData != null)
                    {
                        // 读取属性值
                        object? propValue = cellProp.GetValue(targetData);
                        if (propValue != null && propValue is float)
                        {
                            float value = (float)propValue;
                            cellValue = value.ToString("0.00");
                        }
                        // 兼容可空float类型(float?)
                        else if (propValue != null && propValue is float?)
                        {
                            float? nullableValue = (float?)propValue;
                            if (nullableValue.HasValue)
                            {
                                cellValue = nullableValue.Value.ToString("0.00");
                            }
                        }
                    }
                    // 赋值到Cells字典(即使属性不存在，也会赋值为空字符串)
                    hourData.Cells[cellKey] = cellValue;
                }

                return hourData;
            }).ToList();

            // 异步返回结果(适配async方法)
            return hourDataList;
        }

    }
}
