
using System.Text.Json.Serialization;

namespace CenterBackend.Dto;

/// <summary>
/// 折线图数据总DTO(对应前端LineChartData接口)
/// </summary>
public class LineChartDataDto
{
    /// <summary>
    /// X轴标签数组(如：["01日", "02日"]、["第1周", "第2周"])
    /// </summary>
    [JsonPropertyName("xAxis")] // 确保序列化后字段名与前端一致
    public string[] XAxis { get; set; } = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];

    /// <summary>
    /// 折线数据系列(支持单条/多条折线)
    /// </summary>
    [JsonPropertyName("series")]
    public List<LineChartSeriesDto> Series { get; set; } = [];
}

/// <summary>
/// 折线图单条系列数据DTO
/// </summary>
public class LineChartSeriesDto
{
    /// <summary>
    /// 折线名称(图例显示，如："日产量"、"实际产量")
    /// </summary>
    [JsonPropertyName("name")]
    public string Name { get; set; } = "default";

    /// <summary>
    /// 折线数值数组(与XAxis长度一致)
    /// </summary>
    [JsonPropertyName("data")]
    public float?[] Data { get; set; } = [0, 0, 0];
}