namespace CenterBackend.Models.ExcelDataView
{
    public class BaseSheet
    {
        public required SheetType SheetType { get; set; } = SheetType.OtherReport;
        public required string ReportedTime { get; set; } 
        public required string Directory { get; set; }
        public required string FileName { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }


    public enum SheetType
    {
        DayReport,
        MonthReport,
        YearReport,
        WeekReport,
        OtherReport
    }
}
