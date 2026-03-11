using CenterBackend.Dto;
using CenterBackend.IServices;
using CenterBackend.Services;
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
                if (type <= _mockHeaderMains.Count && type >= 0 ) 
                {
                    return Ok(_mockHeaderMains[type - 1]); // 返回200 
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
        public async Task<ActionResult<List<HourDataDto>>> GetHourData(string date, string type)
        {
            // 1. 校验日期格式
            //if (!DateTime.TryParseExact(getHourDatasDto.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var queryDate))
            //{
            //    return BadRequest(new { message = "日期格式错误，请传入YYYY-MM-DD格式" });
            //}
            try
            {

                var resultList = await _reportRecordService.getHourDataTableOne(date, type);

                // 直接返回结果列表
                return Ok(resultList);
            }
            catch (Exception ex)
            {
                // 生产环境建议添加日志记录
                // _logger.LogError(ex, "查询小时数据失败，日期：{QueryDate}", getHourDatasDto.Date);
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


        private readonly List<TableHeaderMainDto> _mockHeaderMains = new()
{
    // 表1
    new TableHeaderMainDto
    {
        // 分组表头：将检测指标归为一个分组
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell1", "Cell2", "Cell3", "Cell4", "Cell5" },
                Label = "闪发器冷凝液检测数",
                Unit = ""
            }
        },
        // 普通表头：与原结构一致
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto { Prop = "Cell1", Label = "COD(mg/L)", Unit = "mg/L" },
            new TableHeaderDto { Prop = "Cell2", Label = "TCN/总腈(mg/L)", Unit = "mg/L" },
            new TableHeaderDto { Prop = "Cell3", Label = "NH3-N氨氮(mg/L)", Unit = "mg/L" },
            new TableHeaderDto { Prop = "Cell4", Label = "HCHO甲醛(mg/L)", Unit = "mg/L" },
            new TableHeaderDto { Prop = "Cell5", Label = "闪发器冷凝液pH", Unit = "" },
        }
    },
 

   // 表2
    new TableHeaderMainDto
    {
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell11", "Cell12", "Cell13", "Cell14", "Cell15", "Cell16" ,"Cell17", "Cell18"},
                Label = "反应液检测数据记录表",
                Unit = ""
            },
        },
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto { Prop = "Cell11", Label = "二乙腈含量-化分（%）", Unit = "%" },
            new TableHeaderDto { Prop = "Cell12", Label = "二乙腈含量-色谱（%）", Unit = "%" },
            new TableHeaderDto { Prop = "Cell13", Label = "羟基乙腈残余（%）", Unit = "%" },
            new TableHeaderDto { Prop = "Cell14", Label = "羟基乙腈残余（g/L）", Unit = "g/L" },
            new TableHeaderDto { Prop = "Cell15", Label = "甘氨腈（g/L）", Unit = "g/L" },
            new TableHeaderDto { Prop = "Cell16", Label = "三乙腈（g/L）", Unit = "g/L" },
            new TableHeaderDto { Prop = "Cell17", Label = "反应液检测数据pH", Unit = "" },
            new TableHeaderDto { Prop = "Cell18", Label = "反应液检测数据pH", Unit = "" },
        }
    },


    // 表3
    new TableHeaderMainDto
    {
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell21", "Cell22", "Cell23", "Cell24", "Cell25", "Cell26" },
                Label = "一次结晶物/一次产品",
                Unit = ""
            },
            new GroupHeader
            {
                Props = new List<string> { "Cell31", "Cell32", "Cell33", "Cell34", "Cell35", "Cell36", "Cell37" },
                Label = "二次结晶物/二次产品",
                Unit = ""
            }
        },
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto { Prop = "Cell21", Label = "二乙腈含量-化分 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell22", Label = "二乙腈含量-色谱 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell23", Label = "水分含量 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell24", Label = "二乙腈 + 水 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell25", Label = "未知物含量 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell26", Label = "产量 (kg)", Unit = "kg" },
            new TableHeaderDto { Prop = "Cell31", Label = "二乙腈含量-化分 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell32", Label = "二乙腈含量-色谱 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell33", Label = "水分含量 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell34", Label = "二乙腈 + 水 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell35", Label = "未知物含量 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell36", Label = "产量 (kg)", Unit = "kg" },
            new TableHeaderDto { Prop = "Cell37", Label = "折百产量 (kg)", Unit = "kg" },
        }
    },
  

    // 表4
    new TableHeaderMainDto
    {
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell41", "Cell42", "Cell43", "Cell44", "Cell45" },
                Label = "一次母液成分检测",
                Unit = ""
            }
        },
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto { Prop = "Cell41", Label = "二乙腈含量-化分 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell42", Label = "二乙腈含量-色谱 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell43", Label = "羟基乙睛残余-化分 (%)", Unit = "%" },
            new TableHeaderDto { Prop = "Cell44", Label = "羟基乙睛残余-色谱 (g/L)", Unit = "g/L" },
            new TableHeaderDto { Prop = "Cell45", Label = "硫铵 (g/L)", Unit = "g/L" },
        }
    },
  

  // 表5
    new TableHeaderMainDto
    {
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell51", "Cell52", "Cell53", "Cell54", "Cell55", "Cell56" },
                Label = "母液脱色前检测数据",
                Unit = ""
            },
            new GroupHeader
            {
                Props = new List<string> { "Cell57", "Cell58", "Cell59", "Cell60", "Cell61", "Cell62" },
                Label = "母液脱色后检测数据",
                Unit = ""
            },
            new GroupHeader
            {
                Props = new List<string> { "Cell63", "Cell64" },
                Label = "废液/耗材消耗",
                Unit = ""
            }
        },
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto {Prop =  "Cell51", Label = "二乙腈含量 (化分（%）)" },
            new TableHeaderDto { Prop = "Cell52", Label = "二乙腈含量 (色谱（%）)" },
            new TableHeaderDto { Prop = "Cell53", Label = "羟基乙腈残余 (%)" },
            new TableHeaderDto { Prop = "Cell54", Label = "羟基乙腈残余 (g/L)" },
            new TableHeaderDto { Prop = "Cell55", Label = "硫铵 (g/L)" },
            new TableHeaderDto { Prop = "Cell56", Label = "透光率 (%)" },
            new TableHeaderDto { Prop = "Cell57", Label = "二乙腈含量 (化分（%）)" },
            new TableHeaderDto { Prop = "Cell58", Label = "二乙腈含量 (色谱（%）)" },
            new TableHeaderDto { Prop = "Cell59", Label = "羟基乙腈残余 (%)" },
            new TableHeaderDto { Prop = "Cell60", Label = "羟基乙腈残余 (g/L)" },
            new TableHeaderDto { Prop = "Cell61", Label = "硫铵 (g/L)" },
            new TableHeaderDto { Prop = "Cell62", Label = "透光率 (%)" },
            new TableHeaderDto { Prop = "Cell63", Label = "废液中二乙睛含量 %" },
            new TableHeaderDto { Prop = "Cell64", Label = "活性炭消耗 (kg)" },
        }
    },
   

   // 表6
    new TableHeaderMainDto
    {
        groupHeaders = new List<GroupHeader>
        {
            new GroupHeader
            {
                Props = new List<string> { "Cell71", "Cell72", "Cell73" },
                Label = "公用工程消耗",
                Unit = ""
            }
        },
        tableHeaders = new List<TableHeaderDto>
        {
            new TableHeaderDto { Prop = "hour", Label = "小时", Unit = "" },
            new TableHeaderDto { Prop = "Cell71", Label = "蒸汽总消耗（t）" },
            new TableHeaderDto { Prop = "Cell72", Label = "脱盐水消耗（t）" },
            new TableHeaderDto { Prop = "Cell73", Label = "电消耗（KWh）" },
        }
    }
    
};

    }
}
