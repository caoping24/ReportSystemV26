using CenterBackend.common;
using CenterBackend.Dto;
using CenterBackend.IServices;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DashboardController : ControllerBase
    {
        private readonly IDashboardService _dashboardService;

        public DashboardController(IDashboardService _dashboardService)
        {
            this._dashboardService = _dashboardService;
        }

        /// <summary>
        /// 分页记录列表
        /// </summary>
        /// <param name="request">分页参数</param>
        /// <returns>分页结果</returns>
        [HttpGet("getLineChartOne")]
        public async Task<BaseResponse<LineChartDataDto>> GetLineChartOne()
        {
            try
            {

                var result = await _dashboardService.getLineChartOne(DateTime.Now);
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<LineChartDataDto>.error();
            }
        }

        [HttpGet("getLineChartTwo")]
        public async Task<BaseResponse<LineChartDataDto>> GetLineChartTwo()
        {
            try
            {
                var result = await _dashboardService.getLineCharTwo(DateTime.Now);
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<LineChartDataDto>.error();
            }
        }

        [HttpGet("getLineChartThree")]
        public async Task<BaseResponse<LineChartDataDto>> GetLineChartThree()
        {
            try
            {
                var result = await _dashboardService.getLineCharThree(DateTime.Now);
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<LineChartDataDto>.error();
            }
        }


        [HttpGet("getPieChart")]
        public async Task<BaseResponse<List<PieChartItemDto>>> GetPieChart()
        {
            try
            {

                var result = await _dashboardService.getPieChart(DateTime.Now);
                return ResultUtils<List<PieChartItemDto>>.Success(result);
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<List<PieChartItemDto>>.error();
            }
        }


        [HttpGet("getCoreChart")]
        public async Task<BaseResponse<CoreChartDto>> GetCoreChart()
        {
            try
            {

                var result = await _dashboardService.getCoreChart(DateTime.Now);
                return ResultUtils<CoreChartDto>.Success(result);
            }
            catch (Exception ex)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<CoreChartDto>.error();
            }
        }

    }
}
