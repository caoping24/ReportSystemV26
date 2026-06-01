
using CenterBackend.Dto;
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
        Task<bool> UpdateSourceDataFieldAsync(string dateStr, int hour, string prop, string valueStr);

        Task<List<HourDataDto>> getHourDataTableOne(String date, String type);

    }
}
