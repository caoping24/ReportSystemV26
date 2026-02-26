namespace CenterReport.Repository.Utils
{
    // 分页请求参数
    public class PaginationRequest
    {
        /// <summary>
        /// 当前页码（默认第1页）
        /// </summary>
        public int PageIndex { get; set; } = 1;

        /// <summary>
        /// 每页条数（默认10条）
        /// </summary>
        public int PageSize { get; set; } = 10;

        public int Type { get; set; } = 1;
    }

    // 分页响应结果
    public class PaginationResult<T>
    {
        /// <summary>
        /// 当前页码
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// 每页条数
        /// </summary>
        public int PageSize { get; set; }

        /// <summary>
        /// 总记录数
        /// </summary>
        public long TotalCount { get; set; }

        /// <summary>
        /// 总页数
        /// </summary>
        public int TotalPages => (int)Math.Ceiling((double)TotalCount / PageSize);

        /// <summary>
        /// 当前页数据列表
        /// </summary>
        public List<T> Data { get; set; } = new List<T>();
    }
}
