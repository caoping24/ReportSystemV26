///
///2026年7月8日 新增 用于删选数据
///
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace CenterBackend.Models.Filters
{
    /// <summary>
    /// 字段范围筛选配置 — 值必须在 [MinValue, MaxValue] 内才保留
    /// </summary>
    [Table("FieldRangeFilters")]
    public class FieldRangeFilter
    {
        [Key]
        public int Id { get; set; }

        /// <summary>字段名，如 "Cell1" ~ "Cell150"</summary>
        [Required, MaxLength(128)]
        public string FieldName { get; set; } = string.Empty;
        /// <summary>注释说明</summary>
        [MaxLength(500)]
        public string? Comment { get; set; }
        /// <summary>最小值（null = 不检查下限）</summary>
        public float? MinValue { get; set; }

        /// <summary>最大值（null = 不检查上限）</summary>
        public float? MaxValue { get; set; }

        public bool IsActive { get; set; } = true;

        /// <summary>
        /// 判断原始值是否在 [MinValue, MaxValue] 范围内
        /// </summary>
        public bool IsValid(float? value)
        {
            if (!value.HasValue)
                return false;

            if (MinValue.HasValue && value.Value < MinValue.Value)
                return false;
            if (MaxValue.HasValue && value.Value > MaxValue.Value)
                return false;

            return true;
        }
    }
}