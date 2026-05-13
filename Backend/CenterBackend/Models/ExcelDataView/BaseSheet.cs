using System.Globalization;

namespace CenterBackend.Models.ExcelDataView
{
    public class BaseSheet
    {
        public required SheetType SheetType { get; set; } = SheetType.OtherReport;
        public required DateTime ReportedTime { get; set; }
        public required string Directory { get; set; }
        public required string FileName { get; set; }
        public required string ModFilePath { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string WeekNumberInYear
        {
            get
            {
                Calendar calendar = CultureInfo.InvariantCulture.Calendar;// 计算是当年第几周
                int weekOfYear = calendar.GetWeekOfYear(
                    this.ReportedTime.Date,
                    CalendarWeekRule.FirstDay,    // 周规则
                    DayOfWeek.Monday              // 一周起始日(周一)
                );
                var temp = $"{ReportedTime.Year}年{weekOfYear}周";
                return temp;
            }
        }
        public string ReportDate
        {
            get => ReportedTime.ToString("yyyy-MM-dd");
        }
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
