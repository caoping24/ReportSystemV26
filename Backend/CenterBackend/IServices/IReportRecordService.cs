
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;


namespace CenterBackend.IReportServices
{
    public interface IReportRecordService
    {
        /// <summary>
        /// 分页查询报表
        /// </summary>
        /// <param name="request">分页参数</param>
        /// <returns>分页结果</returns>
        Task<PaginationResult<ReportRecord>> GetReportsByPageAsync(PaginationRequest request);

    }
}
