using CenterBackend.IReportServices;
using CenterReport.Repository;
using CenterReport.Repository.Models;
using CenterReport.Repository.Utils;

namespace CenterBackend.Services
{
    public class ReportRecordService : IReportRecordService
    {
        private readonly IReportRecordRepository<ReportRecord> _reportRecord;


        public ReportRecordService(IReportRecordRepository<ReportRecord> reportRecord)
        {
            this._reportRecord = reportRecord;
        }


        public async Task<PaginationResult<ReportRecord>> GetReportsByPageAsync(PaginationRequest request)
        {
            return await _reportRecord.GetReportByPageAsync(request);
        }

    }


}
