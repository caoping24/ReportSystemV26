using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterReport.Repository.IServices;
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

        public ReportRecordController(IReportRecordService reportRecordService)
        {
            this._reportRecordService = reportRecordService;
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


        [HttpGet("Headers")]
        public async Task<ActionResult<List<TableHeaderDto>>> GetHeaders(int type)
        {
            try
            {
                if (type <= _mockHeaders.Count && type >= 0 ) 
                {
                    return Ok(_mockHeaders[type - 1]); // 返回200 
                }
                else
                {
                    return BadRequest(new { message = "传入Type不合法"});
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "查询失败", detail = ex.Message });
            }
        }

        [HttpGet("HourData")]
        public async Task<ActionResult<List<HourDataDto>>> GetHourData(GetHourDatasDto getHourDatasDto)
        {
            // 1. 校验日期格式
            if (!DateTime.TryParseExact(getHourDatasDto.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var queryDate))
            {
                return BadRequest(new { message = "日期格式错误，请传入YYYY-MM-DD格式" });
            }

            try
            {

                
                var hourList = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23 };
                var hourData = new HourDataDto
                {
                    Hour = 1,
                    Date = getHourDatasDto.date,
                    IsNextDay = false ,//isNextDay, // 赋值修正后的禁用标识
                    Cells = new Dictionary<string, string>() // 确保Cells初始化，避免空引用
                };
                return new List<hourData>[];
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
                await _reportRecordService.UpdateSourceDataFieldAsync(
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

        private readonly List<List<TableHeaderDto>> _mockHeaders = new()
        {
            new List<TableHeaderDto>//表1-闪发器冷凝液检测数据
            {
                new TableHeaderDto { Prop = "hour", Label = "小时" },
                
                new TableHeaderDto { Prop = "Cell151", Label = "COD(mg/L)"  },
                new TableHeaderDto { Prop = "Cell152", Label = "TCN/总腈(mg/L)" },
                new TableHeaderDto { Prop = "Cell153", Label = "NH3-N氨氮(mg/L)" },
                new TableHeaderDto { Prop = "Cell154", Label = "HCHO甲醛(mg/L)" },
                new TableHeaderDto { Prop = "Cell155", Label = "闪发器冷凝液ph" },
            },
            new List<TableHeaderDto>//表2-反应液检测数据
            {
                new TableHeaderDto { Prop = "hour", Label = "小时" },
                new TableHeaderDto { Prop = "Cell161", Label = "二乙腈含量-化分（%）" },
                new TableHeaderDto { Prop = "Cell162", Label = "二乙腈含量-色谱（%）" },
                new TableHeaderDto { Prop = "Cell163", Label = "羟基乙腈残余（%）" },
                new TableHeaderDto { Prop = "Cell164", Label = "羟基乙腈残余（g/L）" },
                new TableHeaderDto { Prop = "Cell165", Label = "甘氨腈（g/L）" },
                new TableHeaderDto { Prop = "Cell166", Label = "三乙腈（g/L）" },
                new TableHeaderDto { Prop = "Cell167", Label = "反应液检测数据pH" },
                new TableHeaderDto { Prop = "Cell168", Label = "反应液检测数据pH" },
            },
            new List<TableHeaderDto>//表3-结晶检测数据
            {
                new TableHeaderDto { Prop = "Cell171", Label = "二乙腈含量-化分 (%)" },
                new TableHeaderDto { Prop = "Cell172", Label = "二乙腈含量-色谱 (%)" },
                new TableHeaderDto { Prop = "Cell173", Label = "水分含量 (%)" },
                new TableHeaderDto { Prop = "Cell174", Label = "二乙腈 + 水 (%)" },
                new TableHeaderDto { Prop = "Cell175", Label = "未知物含量 (%)" },
                new TableHeaderDto { Prop = "Cell176", Label = "产量 (kg)" },
                new TableHeaderDto { Prop = "Cell177", Label = "折百产量 (kg)" },
                new TableHeaderDto { Prop = "Cell181", Label = "二乙腈含量-化分 (%)" },
                new TableHeaderDto { Prop = "Cell182", Label = "二乙腈含量-色谱 (%)" },
                new TableHeaderDto { Prop = "Cell183", Label = "水分含量 (%)" },
                new TableHeaderDto { Prop = "Cell184", Label = "二乙腈 + 水 (%)" },
                new TableHeaderDto { Prop = "Cell185", Label = "未知物含量 (%)" },
                new TableHeaderDto { Prop = "Cell186", Label = "产量 (kg)" },
                new TableHeaderDto { Prop = "Cell187", Label = "折百产量 (kg)" },
                new TableHeaderDto { Prop = "Cell191", Label = "二乙腈含量-化分 (%)" },
                new TableHeaderDto { Prop = "Cell192", Label = "二乙腈含量-色谱 (%)" },
                new TableHeaderDto { Prop = "Cell193", Label = "水分含量 (%)" },
                new TableHeaderDto { Prop = "Cell194", Label = "二乙腈 + 水 (%)" },
                new TableHeaderDto { Prop = "Cell195", Label = "未知物含量 (%)" },
                new TableHeaderDto { Prop = "Cell196", Label = "产量 (kg)" },
                new TableHeaderDto { Prop = "Cell197", Label = "折百产量 (kg)" },
                new TableHeaderDto { Prop = "Cell198", Label = "班产 (kg)" },
            },
            new List<TableHeaderDto>//表4-一次母液分析数据
            {
                new TableHeaderDto { Prop = "Cell201", Label = "二乙腈含量-化分 (%)" },
                new TableHeaderDto { Prop = "Cell202", Label = "二乙腈含量-色谱 (%)" },
                new TableHeaderDto { Prop = "Cell203", Label = "羟基乙睛残余-化分 (%)" },
                new TableHeaderDto { Prop = "Cell204", Label = "羟基乙睛残余-色谱 (g/L)" },
                new TableHeaderDto { Prop = "Cell205", Label = "硫铵 (g/L)" },
            },
            new List<TableHeaderDto>//表5-母液驼色检测数据
            {
                new TableHeaderDto {Prop =  "Cell211", Label = "二乙腈含量 (化分（%）)" },
                new TableHeaderDto { Prop = "Cell212", Label = "二乙腈含量 (色谱（%）)" },
                new TableHeaderDto { Prop = "Cell213", Label = "羟基乙腈残余 (%)" },
                new TableHeaderDto { Prop = "Cell214", Label = "羟基乙腈残余 (g/L)" },
                new TableHeaderDto { Prop = "Cell215", Label = "硫铵 (g/L)" },
                new TableHeaderDto { Prop = "Cell216", Label = "透光率 (%)" },
                new TableHeaderDto { Prop = "Cell221", Label = "二乙腈含量 (化分（%）)" },
                new TableHeaderDto { Prop = "Cell222", Label = "二乙腈含量 (色谱（%）)" },
                new TableHeaderDto { Prop = "Cell223", Label = "羟基乙腈残余 (%)" },
                new TableHeaderDto { Prop = "Cell224", Label = "羟基乙腈残余 (g/L)" },
                new TableHeaderDto { Prop = "Cell225", Label = "硫铵 (g/L)" },
                new TableHeaderDto { Prop = "Cell226", Label = "透光率 (%)" },
                new TableHeaderDto { Prop = "Cell227", Label = "活性炭消耗 (kg)" },
            },
            new List<TableHeaderDto>//表6-日常消耗
            {
                new TableHeaderDto { Prop = "hour", Label = "小时" },
                new TableHeaderDto { Prop = "Cell230", Label = "蒸汽总消耗（t）" },
                new TableHeaderDto { Prop = "Cell231", Label = "脱盐水消耗（t）" },
                new TableHeaderDto { Prop = "Cell232", Label = "电消耗（KWh）" },


            },
        };
    }
}
