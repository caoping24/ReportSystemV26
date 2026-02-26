using System.ComponentModel.DataAnnotations.Schema;
using System;

namespace CenterReport.Repository.Models
{
    [Table("SourceData")]
    public class SourceData
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
        public float? Cell51 { get; set; }
        public float? Cell52 { get; set; }
        public float? Cell53 { get; set; }
        public float? Cell54 { get; set; }
        public float? Cell55 { get; set; }
        public float? Cell56 { get; set; }
        public float? Cell57 { get; set; }
        public float? Cell58 { get; set; }
        public float? Cell59 { get; set; }
        public float? Cell60 { get; set; }
        public float? Cell61 { get; set; }
        public float? Cell62 { get; set; }
        public float? Cell63 { get; set; }
        public float? Cell64 { get; set; }
        public float? Cell65 { get; set; }
        public float? Cell66 { get; set; }
        public float? Cell67 { get; set; }
        public float? Cell68 { get; set; }
        public float? Cell69 { get; set; }
        public float? Cell70 { get; set; }
        public float? Cell71 { get; set; }
        public float? Cell72 { get; set; }
        public float? Cell73 { get; set; }
        public float? Cell74 { get; set; }
        public float? Cell75 { get; set; }
        public float? Cell76 { get; set; }
        public float? Cell77 { get; set; }
        public float? Cell78 { get; set; }
        public float? Cell79 { get; set; }
        public float? Cell80 { get; set; }
        public float? Cell81 { get; set; }
        public float? Cell82 { get; set; }
        public float? Cell83 { get; set; }
        public float? Cell84 { get; set; }
        public float? Cell85 { get; set; }
        public float? Cell86 { get; set; }
        public float? Cell87 { get; set; }
        public float? Cell88 { get; set; }
        public float? Cell89 { get; set; }
        public float? Cell90 { get; set; }
        public float? Cell91 { get; set; }
        public float? Cell92 { get; set; }
        public float? Cell93 { get; set; }
        public float? Cell94 { get; set; }
        public float? Cell95 { get; set; }
        public float? Cell96 { get; set; }
        public float? Cell97 { get; set; }
        public float? Cell98 { get; set; }
        public float? Cell99 { get; set; }
        public float? Cell100 { get; set; }
        public float? Cell101 { get; set; }
        public float? Cell102 { get; set; }
        public float? Cell103 { get; set; }
        public float? Cell104 { get; set; }
        public float? Cell105 { get; set; }
        public float? Cell106 { get; set; }
        public float? Cell107 { get; set; }
        public float? Cell108 { get; set; }
        public float? Cell109 { get; set; }
        public float? Cell110 { get; set; }
        public float? Cell111 { get; set; }
        public float? Cell112 { get; set; }
        public float? Cell113 { get; set; }
        public float? Cell114 { get; set; }
        public float? Cell115 { get; set; }
        public float? Cell116 { get; set; }
        public float? Cell117 { get; set; }
        public float? Cell118 { get; set; }
        public float? Cell119 { get; set; }
        public float? Cell120 { get; set; }
        public float? Cell121 { get; set; }
        public float? Cell122 { get; set; }
        public float? Cell123 { get; set; }
        public float? Cell124 { get; set; }
        public float? Cell125 { get; set; }
        public float? Cell126 { get; set; }
        public float? Cell127 { get; set; }
        public float? Cell128 { get; set; }
        public float? Cell129 { get; set; }
        public float? Cell130 { get; set; }
        public float? Cell131 { get; set; }
        public float? Cell132 { get; set; }
        public float? Cell133 { get; set; }
        public float? Cell134 { get; set; }
        public float? Cell135 { get; set; }
        public float? Cell136 { get; set; }
        public float? Cell137 { get; set; }
        public float? Cell138 { get; set; }
        public float? Cell139 { get; set; }
        public float? Cell140 { get; set; }
        public float? Cell141 { get; set; }
        public float? Cell142 { get; set; }
        public float? Cell143 { get; set; }
        public float? Cell144 { get; set; }
        public float? Cell145 { get; set; }
        public float? Cell146 { get; set; }
        public float? Cell147 { get; set; }
        public float? Cell148 { get; set; }
        public float? Cell149 { get; set; }
        public float? Cell150 { get; set; }
    }
}