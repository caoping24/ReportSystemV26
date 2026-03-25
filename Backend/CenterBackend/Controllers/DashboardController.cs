using CenterBackend.common;
using CenterBackend.Dto;
using CenterBackend.IServices;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DashboardController(IDashboardService _dashboardService) : ControllerBase
    {
        private readonly IDashboardService _dashboardService = _dashboardService;
        //第一页卡片1
        [HttpGet("GetPage1CoreChart1")]
        public async Task<BaseResponse<CoreChartDto>> GetPage1CoreChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage1CoreChart1();
                return ResultUtils<CoreChartDto>.Success(result);
            }
            catch (Exception)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<CoreChartDto>.Error();
            }
        }
        //第一页曲线1
        [HttpGet("GetPage1LineChart1")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage1LineChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage1LineChart1();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第一页曲线2
        [HttpGet("GetPage1LineChart2")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage1LineChart2()
        {
            try
            {
                var result = await _dashboardService.GetPage1LineChart2();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第一页曲线3
        [HttpGet("GetPage1LineChart3")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage1LineChart3()
        {
            try
            {
                var result = await _dashboardService.GetPage1LineChart3();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第一页曲线4
        [HttpGet("GetPage1LineChart4")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage1LineChart4()
        {
            try
            {
                var result = await _dashboardService.GetPage1LineChart4();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第一页曲线5
        [HttpGet("GetPage1LineChart5")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage1LineChart5()
        {
            try
            {
                var result = await _dashboardService.GetPage1LineChart5();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第二页卡片1
        [HttpGet("GetPage2CoreChart1")]
        public async Task<BaseResponse<CoreChartDto>> GetPage2CoreChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage2CoreChart1();
                return ResultUtils<CoreChartDto>.Success(result);
            }
            catch (Exception)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<CoreChartDto>.Error();
            }
        }
        //第二页曲线1
        [HttpGet("GetPage2LineChart1")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage2LineChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage2LineChart1();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第二页曲线2
        [HttpGet("GetPage2LineChart2")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage2LineChart2()
        {
            try
            {
                var result = await _dashboardService.GetPage2LineChart2();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第二页曲线3
        [HttpGet("GetPage2LineChart3")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage2LineChart3()
        {
            try
            {
                var result = await _dashboardService.GetPage2LineChart3();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第二页曲线4
        [HttpGet("GetPage2LineChart4")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage2LineChart4()
        {
            try
            {
                var result = await _dashboardService.GetPage2LineChart4();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第二页曲线5
        [HttpGet("GetPage2LineChart5")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage2LineChart5()
        {
            try
            {
                var result = await _dashboardService.GetPage2LineChart5();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第三页卡片1
        [HttpGet("GetPage3CoreChart1")]
        public async Task<BaseResponse<CoreChartDto>> GetPage3CoreChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage3CoreChart1();
                return ResultUtils<CoreChartDto>.Success(result);
            }
            catch (Exception)
            {
                // 异常处理（实际项目可封装全局异常过滤器）
                return ResultUtils<CoreChartDto>.Error();
            }
        }
        //第三页曲线1
        [HttpGet("GetPage3LineChart1")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage3LineChart1()
        {
            try
            {
                var result = await _dashboardService.GetPage3LineChart1();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第三页曲线2
        [HttpGet("GetPage3LineChart2")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage3LineChart2()
        {
            try
            {
                var result = await _dashboardService.GetPage3LineChart2();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第三页曲线3
        [HttpGet("GetPage3LineChart3")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage3LineChart3()
        {
            try
            {
                var result = await _dashboardService.GetPage3LineChart3();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }
        //第三页曲线4
        [HttpGet("GetPage3LineChart4")]
        public async Task<BaseResponse<LineChartDataDto>> GetPage3LineChart4()
        {
            try
            {
                var result = await _dashboardService.GetPage3LineChart4();
                return ResultUtils<LineChartDataDto>.Success(result);
            }
            catch (Exception)
            {
                return ResultUtils<LineChartDataDto>.Error();
            }
        }

    }
}
