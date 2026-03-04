using CenterBackend.Models.ExcelDataView;

namespace CenterBackend.Models.CalculateData
{
    public class ReportInfo
    {
        public SheetType SheetType { get; set; }
        public DateTime TimeStart { get; set; }
        public DateTime TimeEnd { get; set; }
    }
}
