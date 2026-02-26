namespace CenterBackend.Dto
{
    public class CoreChartDto
    {
        /// <summary>
        /// 昨日数据
        /// </summary>
        public float Yesterday { get; set; }

        /// <summary>
        /// 本周数据
        /// </summary>
        public float Week { get; set; }

        /// <summary>
        /// 本月数据
        /// </summary>
        public float Month { get; set; }

        /// <summary>
        /// 本年数据
        /// </summary>
        public float Year { get; set; }
    }
}
