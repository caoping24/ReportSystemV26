using CenterBackend.Dto;
using CenterBackend.Models;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;

namespace CenterBackend.IServices
{
    public interface IReportService
    {
        Task<bool> RebuildReport(PathAndName fileInfo);
    }
}
