using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace CenterReport.Repository.Models
{
    [Table("OperatorInputData")]
    public class OperatorInputData
    {
        /// <summary>
        /// 主键ID（自增）
        /// </summary>
        [Key]
        public long Id { get; set; }

        /// <summary>
        /// 上报时间（对应SQL中的ReportedTime）
        /// </summary>
        public DateTime ReportedTime { get; set; }

        /// <summary>
        /// 最后修改时间（对应SQL中的LastChange，默认值由数据库设置）
        /// </summary>
        public DateTime LastChange { get; set; } = DateTime.Now;

        // 类型（可空）
        public int? Type { get; set; } = 0;

        // PH值（可空）
        public int? PH { get; set; }

        // Cell1 到 Cell150 字段（对应SQL中的real类型，映射为float）
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
        public float? Cell6 { get; set; }
        public float? Cell7 { get; set; }
        public float? Cell8 { get; set; }
        public float? Cell9 { get; set; }
        public float? Cell10 { get; set; }
        public float? Cell11 { get; set; }
        public float? Cell12 { get; set; }
        public float? Cell13 { get; set; }
        public float? Cell14 { get; set; }
        public float? Cell15 { get; set; }
        public float? Cell16 { get; set; }
        public float? Cell17 { get; set; }
        public float? Cell18 { get; set; }
        public float? Cell19 { get; set; }
        public float? Cell20 { get; set; }
        public float? Cell21 { get; set; }
        public float? Cell22 { get; set; }
        public float? Cell23 { get; set; }
        public float? Cell24 { get; set; }
        public float? Cell25 { get; set; }
        public float? Cell26 { get; set; }
        public float? Cell27 { get; set; }
        public float? Cell28 { get; set; }
        public float? Cell29 { get; set; }
        public float? Cell30 { get; set; }
        public float? Cell31 { get; set; }
        public float? Cell32 { get; set; }
        public float? Cell33 { get; set; }
        public float? Cell34 { get; set; }
        public float? Cell35 { get; set; }
        public float? Cell36 { get; set; }
        public float? Cell37 { get; set; }
        public float? Cell38 { get; set; }
        public float? Cell39 { get; set; }
        public float? Cell40 { get; set; }
        public float? Cell41 { get; set; }
        public float? Cell42 { get; set; }
        public float? Cell43 { get; set; }
        public float? Cell44 { get; set; }
        public float? Cell45 { get; set; }
        public float? Cell46 { get; set; }
        public float? Cell47 { get; set; }
        public float? Cell48 { get; set; }
        public float? Cell49 { get; set; }
        public float? Cell50 { get; set; }
        
    }
}