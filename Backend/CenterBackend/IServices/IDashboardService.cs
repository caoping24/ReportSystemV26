using CenterBackend.Dto;

namespace CenterBackend.IServices
{
    public interface IDashboardService
    {
        Task<CoreChartDto> GetPage1CoreChart1();
        Task<LineChartDataDto> GetPage1LineChart1();
        Task<LineChartDataDto> GetPage1LineChart2();
        Task<LineChartDataDto> GetPage1LineChart3();
        Task<LineChartDataDto> GetPage1LineChart4();
        Task<LineChartDataDto> GetPage1LineChart5();

        Task<CoreChartDto> GetPage2CoreChart1();
        Task<LineChartDataDto> GetPage2LineChart1();
        Task<LineChartDataDto> GetPage2LineChart2();
        Task<LineChartDataDto> GetPage2LineChart3();
        Task<LineChartDataDto> GetPage2LineChart4();
        Task<LineChartDataDto> GetPage2LineChart5();

        Task<CoreChartDto> GetPage3CoreChart1();
        Task<LineChartDataDto> GetPage3LineChart1();
        Task<LineChartDataDto> GetPage3LineChart2();
        Task<LineChartDataDto> GetPage3LineChart3();
        Task<LineChartDataDto> GetPage3LineChart4();
    }
}
