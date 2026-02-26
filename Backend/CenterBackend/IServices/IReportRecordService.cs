
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;


namespace CenterBackend.IServices
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
