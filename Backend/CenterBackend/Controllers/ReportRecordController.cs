using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportRecordController : ControllerBase
    {
        private readonly IReportRecordService _reportRecordService;
        private readonly IReportService _reportService;


        public ReportRecordController(IReportRecordService reportRecordService, IReportService reportService)
        {
            this._reportRecordService = reportRecordService;
            this._reportService = reportService;
        }

        /// <summary>
        /// 分页记录列表
        /// </summary>
        /// <param name="request">分页参数</param>
        /// <returns>分页结果</returns>
        [HttpGet("GetReportByPage")]
        public async Task<ActionResult<PaginationResult<ReportRecord>>> GetReportByPage([FromQuery] PaginationRequest request)
        {
            try
            {
                var result = await _reportRecordService.GetReportsByPageAsync(request);

                if (result?.Data != null && result.Data.Count != 0)// ReportedTime 均为 "yyyy-MM-dd" 格式，按该日期降序排序（最新在前）
                {
                    result.Data = result.Data
                        .OrderByDescending(r => r.ReportedTime)
                        .ToList();
                }

                return Ok(result); // 返回200 + 分页结果
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return StatusCode(500, new { message = "查询失败", detail = ex.Message });
            }
        }

        private readonly List<TableHeaderDto> _mockHeaders = new()
        {
            
            new TableHeaderDto { Prop = "hour", Label = "小时" },
            //反应液检测数据
            new TableHeaderDto { Prop = "Cell29", Label = "二乙腈含量-化分（%）" },
            new TableHeaderDto { Prop = "Cell30", Label = "二乙腈含量-色谱（%）" },
            new TableHeaderDto { Prop = "Cell31", Label = "羟基乙腈残余（%）" },
            new TableHeaderDto { Prop = "Cell32", Label = "羟基乙腈残余（g/L）" },
            new TableHeaderDto { Prop = "Cell33", Label = "甘氨腈（g/L）" },
            new TableHeaderDto { Prop = "Cell34", Label = "三乙腈（g/L）" },
            new TableHeaderDto { Prop = "Cell35", Label = "反应液检测数据pH" },
            //闪发器冷凝液检测数据记录
            new TableHeaderDto { Prop = "Cell56", Label = "COD(mg/L)" },
            new TableHeaderDto { Prop = "Cell57", Label = "TCN/总腈(mg/L)" },
            new TableHeaderDto { Prop = "Cell58", Label = "NH3-N氨氮(mg/L)" },
            new TableHeaderDto { Prop = "Cell59", Label = "HCHO甲醛(mg/L)" },
            new TableHeaderDto { Prop = "Cell60", Label = "闪发器冷凝液ph" },
            //一次母液分析数据
            new TableHeaderDto { Prop = "Cell82", Label = "一次分离产量（Kg）" },
            new TableHeaderDto { Prop = "Cell83", Label = "二乙腈含量化分（%）" },
            new TableHeaderDto { Prop = "Cell84", Label = "二乙腈含量色谱（%）" },
            new TableHeaderDto { Prop = "Cell85", Label = "羟基乙腈残余（%）" },
            new TableHeaderDto { Prop = "Cell86", Label = "羟基乙腈残余（g/L）" },
            new TableHeaderDto { Prop = "Cell87", Label = "硫铵含量（g/L）" },
            //脱色数据记录
            new TableHeaderDto { Prop = "Cell135", Label = "脱色前透光率(%)" },
            new TableHeaderDto { Prop = "Cell136", Label = "脱色后透光率(%)" },
            //巡检参数记录
            new TableHeaderDto { Prop = "Cell137", Label = "脱色输送泵（mbar）" },
            new TableHeaderDto { Prop = "Cell138", Label = "低蒸出料泵（mbar）" },
            new TableHeaderDto { Prop = "Cell139", Label = "低蒸循环泵（mbar）" },
            new TableHeaderDto { Prop = "Cell140", Label = "一次结晶清洗泵（mbar）" },
            new TableHeaderDto { Prop = "Cell141", Label = "二次结晶清洗泵（mbar）" }

        };
        [HttpGet("Headers")]
        public async Task<ActionResult<List<TableHeaderDto>>> GetHeaders()
        {

            try
            {
                return Ok(_mockHeaders); // 返回200 
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "查询失败", detail = ex.Message });
            }
        }

        [HttpGet("HourData")]
        public async Task<ActionResult<List<HourDataDto>>> GetHourData([FromQuery] string date)
        {
            // 1. 校验日期格式
            if (!DateTime.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var queryDate))
            {
                return BadRequest(new { message = "日期格式错误，请传入YYYY-MM-DD格式" });
            }

            try
            {
                // 当天
                DateTime startTime = queryDate.Date.AddHours(0);   
                DateTime endTime = startTime.AddHours(23).AddMinutes(59);         

                var calculatedDatas = await _reportService.GetSourceData(startTime, endTime);

                // 3.构建【带日期维度的分组键】
                var dataWithKey = calculatedDatas.Select(cd => new
                {
                    Data = cd,
                    GroupKey = cd.ReportedTime >= startTime && cd.ReportedTime < endTime
                        ? cd.ReportedTime.Hour
                        : cd.ReportedTime.Hour + 100
                }).ToList();

                // 4. 按唯一分组键分组（保留原有逻辑）
                var hourGroupDict = dataWithKey
                    .GroupBy(item => item.GroupKey)
                    .ToDictionary(g => g.Key, g => g.FirstOrDefault()?.Data);

                var hourList = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };

                var hourDataList = hourList.Select((hour, index) =>// 6. 构建返回数据（核心修改：计算真实时间+动态判定IsNextDay）
                {

                    DateTime realHourTime;
                    realHourTime = queryDate.Date.AddHours(hour);

          
                    int targetKey = hour;
                    hourGroupDict.TryGetValue(targetKey, out var targetData);


                    // 未来时间：该时段的真实时间 大于 当前系统时间
                    bool isFutureTime = realHourTime > DateTime.Now;

                    // IsNextDay=true（前端禁用）：未来时间 OR 无对应数据
                    // IsNextDay=false（前端可编辑）：过去时间 AND 有对应数据
                    bool isNextDay = isFutureTime || targetData == null;

                    // 初始化返回DTO，赋值核心字段
                    var hourData = new HourDataDto
                    {
                        Hour = hour,
                        Date = date,
                        IsNextDay = isNextDay, // 赋值修正后的禁用标识
                        Cells = new Dictionary<string, string>() // 确保Cells初始化，避免空引用
                    };

                    // 7. 填充Cell字段（保留原有格式化逻辑，无数据为空字符串）
                    hourData.Cells["Cell29"] = targetData?.Cell29?.ToString("0.00") ?? "";
                    hourData.Cells["Cell30"] = targetData?.Cell30?.ToString("0.00") ?? "";
                    hourData.Cells["Cell31"] = targetData?.Cell31?.ToString("0.00") ?? "";
                    hourData.Cells["Cell32"] = targetData?.Cell32?.ToString("0.00") ?? "";
                    hourData.Cells["Cell33"] = targetData?.Cell33?.ToString("0.00") ?? "";
                    hourData.Cells["Cell34"] = targetData?.Cell34?.ToString("0.00") ?? "";
                    hourData.Cells["Cell35"] = targetData?.Cell35?.ToString("0.00") ?? "";
                    hourData.Cells["Cell56"] = targetData?.Cell56?.ToString("0.00") ?? "";
                    hourData.Cells["Cell57"] = targetData?.Cell57?.ToString("0.00") ?? "";
                    hourData.Cells["Cell58"] = targetData?.Cell58?.ToString("0.00") ?? "";
                    hourData.Cells["Cell59"] = targetData?.Cell59?.ToString("0.00") ?? "";
                    hourData.Cells["Cell60"] = targetData?.Cell60?.ToString("0.00") ?? "";
                    hourData.Cells["Cell82"] = targetData?.Cell82?.ToString("0.00") ?? "";
                    hourData.Cells["Cell83"] = targetData?.Cell83?.ToString("0.00") ?? "";
                    hourData.Cells["Cell84"] = targetData?.Cell84?.ToString("0.00") ?? "";
                    hourData.Cells["Cell85"] = targetData?.Cell85?.ToString("0.00") ?? "";
                    hourData.Cells["Cell86"] = targetData?.Cell86?.ToString("0.00") ?? "";
                    hourData.Cells["Cell87"] = targetData?.Cell87?.ToString("0.00") ?? "";
                    hourData.Cells["Cell135"] = targetData?.Cell135?.ToString("0.00") ?? "";
                    hourData.Cells["Cell136"] = targetData?.Cell136?.ToString("0.00") ?? "";
                    hourData.Cells["Cell137"] = targetData?.Cell137?.ToString("0.00") ?? "";
                    hourData.Cells["Cell138"] = targetData?.Cell138?.ToString("0.00") ?? "";
                    hourData.Cells["Cell139"] = targetData?.Cell139?.ToString("0.00") ?? "";
                    hourData.Cells["Cell140"] = targetData?.Cell140?.ToString("0.00") ?? "";
                    hourData.Cells["Cell141"] = targetData?.Cell141?.ToString("0.00") ?? "";

                    return hourData;
                }).ToList();

                return Ok(hourDataList);
            }
            catch (Exception ex)
            {
                // 生产环境建议添加日志记录
                // _logger.LogError(ex, "查询小时数据失败，日期：{QueryDate}", date);
                return StatusCode(500, new { message = "查询失败", detail = ex.Message });
            }
        }


        [HttpPost("SaveCell")]
        public async Task<ActionResult<List<TableHeaderDto>>> SaveCell([FromBody] SaveCellRequestDto request)
        {
            // 校验必填参数
            if (string.IsNullOrEmpty(request.Date)
                || string.IsNullOrEmpty(request.Prop)
                || request.Hour < 0 || request.Hour > 23)
            {
                return StatusCode(500, new { message = "参数不合法" });
            }

            if (string.IsNullOrEmpty(request.Value))
            {
                return StatusCode(200, new { message = "数据为空" });
            }
            try
            {
                await _reportService.UpdateSourceDataFieldAsync(
                        dateStr: request.Date,
                        hour: request.Hour,
                        prop: request.Prop,
                        valueStr: request.Value);

                return Ok(); // 返回200 
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "查询失败", detail = ex.Message });
            }
        }


    }
}
