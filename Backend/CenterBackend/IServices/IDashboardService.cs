using CenterBackend.Dto;

namespace CenterBackend.IServices
{
    public interface IDashboardService
    {
        //第一个折线图
        Task<LineChartDataDto> getLineChartOne(DateTime time);
        //第二个折线图
        Task<LineChartDataDto> getLineCharTwo(DateTime time);
        //第三个折线图
        Task<LineChartDataDto> getLineCharThree(DateTime time);
        //饼图
        Task<List<PieChartItemDto>> getPieChart(DateTime time);
        //三个卡片图
        Task<CoreChartDto> getCoreChart(DateTime time);

    }
}
