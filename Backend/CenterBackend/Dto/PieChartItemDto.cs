using System.ComponentModel.DataAnnotations;

namespace CenterBackend.Dto
{
    public class PieChartItemDto
    {
        /// <summary>
        /// 分类名称（如车间A/外协加工，必填）
        /// </summary>

        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 对应数值（产量/占比等，必填，非负）
        /// </summary>
        [Required(ErrorMessage = "数值不能为空")]
        [Range(0, double.MaxValue, ErrorMessage = "数值必须为非负数")]
        public decimal Value { get; set; }
    }
}
