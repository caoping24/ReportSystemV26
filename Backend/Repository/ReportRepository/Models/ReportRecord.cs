using System.ComponentModel.DataAnnotations.Schema;

namespace CenterReport.Repository.Models
{
    [Table("ReportRecord")]
    public class ReportRecord
    {
        /// <summary>
        /// 主键ID（自增）
        /// </summary>
        public long Id { get; set; }

        /// <summary>
        /// 上报时间（对应SQL中的ReportedTime）
        /// </summary>
        public DateTime ReportedTime { get; set; }

        /// <summary>
        /// 最后修改时间（对应SQL中的LastChange，默认值由数据库设置）
        /// </summary>
        public DateTime LastChange { get; set; }

        // 类型（可空）
        public int Type { get; set; } = 0;
        public string? Description { get; set; }

    }

}

